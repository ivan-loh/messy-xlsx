#!/usr/bin/env python3
"""Collect and validate the exact-SHA native ABI workflow result."""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import time
from collections.abc import Callable, Sequence
from pathlib import Path
from typing import Any

DEFAULT_TIMEOUT_SECONDS = 2 * 60 * 60
POLL_INTERVAL_SECONDS = 30.0
LEDGER_SCHEMA_VERSION = 1
PLATFORMS = (
    "manylinux-x86_64",
    "manylinux-aarch64",
    "musllinux-x86_64",
    "musllinux-aarch64",
    "macos-x86_64",
    "macos-arm64",
    "windows-amd64",
)
PYTHON_ABIS = ("cp311", "cp312", "cp313", "cp314")
REQUIRED_ABI_JOB_NAMES = frozenset(
    {
        *(f"native-wheel / {platform}" for platform in PLATFORMS),
        *(
            f"abi3-smoke / {platform} / {python_abi}"
            for platform in PLATFORMS
            for python_abi in PYTHON_ABIS
        ),
    }
)
_REVISION_PATTERN = re.compile(r"[0-9a-f]{40}")


class RunVerificationError(RuntimeError):
    """Raised when an exact-SHA workflow run cannot be accepted."""


GHJSON = Callable[[Sequence[str]], object]
Monotonic = Callable[[], float]
Sleep = Callable[[float], None]


def _run_gh_json(arguments: Sequence[str]) -> object:
    process = subprocess.run(
        ["gh", *arguments],
        text=True,
        capture_output=True,
        check=False,
    )
    if process.returncode != 0:
        detail = process.stderr.strip() or process.stdout.strip()
        raise RunVerificationError(f"gh {' '.join(arguments)} failed: {detail}")
    try:
        return json.loads(process.stdout)
    except json.JSONDecodeError as exc:
        raise RunVerificationError("gh returned invalid JSON") from exc


def _require_revision(revision: str) -> None:
    if _REVISION_PATTERN.fullmatch(revision) is None:
        raise RunVerificationError("revision must be an exact lowercase forty-character commit SHA")


def _as_mapping(value: object, context: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise RunVerificationError(f"{context} must be a JSON object")
    return value


def _as_sequence(value: object, context: str) -> list[object]:
    if not isinstance(value, list):
        raise RunVerificationError(f"{context} must be a JSON array")
    return value


def _run_database_id(run: dict[str, Any]) -> int:
    database_id = run.get("databaseId")
    if not isinstance(database_id, int) or isinstance(database_id, bool):
        raise RunVerificationError("workflow run has an invalid database ID")
    return database_id


def _wait_or_timeout(
    *,
    monotonic: Monotonic,
    sleep: Sleep,
    deadline: float,
    detail: str,
) -> None:
    now = monotonic()
    if now >= deadline:
        raise RunVerificationError(f"timed out waiting for {detail}")
    sleep(min(POLL_INTERVAL_SECONDS, deadline - now))


def _validate_required_jobs(raw_jobs: object) -> list[dict[str, Any]]:
    jobs = [_as_mapping(job, "workflow job") for job in _as_sequence(raw_jobs, "jobs")]
    jobs_by_name: dict[str, dict[str, Any]] = {}
    for job in jobs:
        name = job.get("name")
        if not isinstance(name, str) or not name:
            raise RunVerificationError("workflow job has an invalid name")
        if name in jobs_by_name:
            raise RunVerificationError(f"duplicate workflow job name: {name}")
        jobs_by_name[name] = job

    missing = sorted(REQUIRED_ABI_JOB_NAMES.difference(jobs_by_name))
    if missing:
        raise RunVerificationError(f"missing required jobs: {', '.join(missing)}")

    for name in sorted(REQUIRED_ABI_JOB_NAMES):
        job = jobs_by_name[name]
        status = job.get("status")
        conclusion = job.get("conclusion")
        if status != "completed" or conclusion != "success":
            raise RunVerificationError(
                f"required job {name!r} is not successful: "
                f"status={status!r}, conclusion={conclusion!r}"
            )

    return [
        {
            "database_id": job.get("databaseId"),
            "name": name,
            "status": job.get("status"),
            "conclusion": job.get("conclusion"),
        }
        for name, job in sorted(jobs_by_name.items())
    ]


def _atomic_write_json(output: Path, payload: dict[str, Any]) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(f".{output.name}.{os.getpid()}.tmp")
    temporary.write_text(
        f"{json.dumps(payload, indent=2, sort_keys=True)}\n",
        encoding="utf-8",
    )
    os.replace(temporary, output)


def collect(
    *,
    revision: str,
    workflow: str,
    output: Path,
    gh_json: GHJSON | None = None,
    monotonic: Monotonic | None = None,
    sleep: Sleep | None = None,
    timeout_seconds: float = DEFAULT_TIMEOUT_SECONDS,
) -> dict[str, Any]:
    """Wait for and validate one exact-SHA successful native ABI workflow run."""

    _require_revision(revision)
    if not workflow:
        raise RunVerificationError("workflow must not be empty")
    if timeout_seconds <= 0:
        raise RunVerificationError("timeout_seconds must be positive")

    run_gh = _run_gh_json if gh_json is None else gh_json
    read_clock = time.monotonic if monotonic is None else monotonic
    wait = time.sleep if sleep is None else sleep
    deadline = read_clock() + timeout_seconds

    while True:
        raw_runs = run_gh(
            [
                "run",
                "list",
                "--commit",
                revision,
                "--workflow",
                workflow,
                "--limit",
                "100",
                "--json",
                "databaseId,headSha,status,conclusion,url",
            ]
        )
        runs = [_as_mapping(run, "workflow run") for run in _as_sequence(raw_runs, "workflow runs")]
        exact_runs = [run for run in runs if run.get("headSha") == revision]
        if not exact_runs:
            _wait_or_timeout(
                monotonic=read_clock,
                sleep=wait,
                deadline=deadline,
                detail=f"{workflow} at {revision}",
            )
            continue

        selected = max(exact_runs, key=_run_database_id)
        status = selected.get("status")
        conclusion = selected.get("conclusion")
        if status != "completed":
            _wait_or_timeout(
                monotonic=read_clock,
                sleep=wait,
                deadline=deadline,
                detail=f"run {_run_database_id(selected)} to complete",
            )
            continue
        if conclusion != "success":
            raise RunVerificationError(
                f"workflow run {_run_database_id(selected)} concluded {conclusion!r}"
            )

        run_id = _run_database_id(selected)
        viewed = _as_mapping(
            run_gh(
                [
                    "run",
                    "view",
                    str(run_id),
                    "--json",
                    "databaseId,headSha,status,conclusion,url,jobs",
                ]
            ),
            "workflow run view",
        )
        if viewed.get("headSha") != revision:
            raise RunVerificationError(
                f"run view SHA {viewed.get('headSha')!r} does not match {revision}"
            )
        if viewed.get("status") != "completed" or viewed.get("conclusion") != "success":
            raise RunVerificationError(
                "run view is not completed successfully: "
                f"status={viewed.get('status')!r}, "
                f"conclusion={viewed.get('conclusion')!r}"
            )
        if _run_database_id(viewed) != run_id:
            raise RunVerificationError("run view database ID changed")

        ledger = {
            "schema_version": LEDGER_SCHEMA_VERSION,
            "revision": revision,
            "workflow": workflow,
            "run": {
                "database_id": run_id,
                "head_sha": revision,
                "status": "completed",
                "conclusion": "success",
                "url": viewed.get("url"),
            },
            "jobs": _validate_required_jobs(viewed.get("jobs")),
        }
        _atomic_write_json(output, ledger)
        return ledger


def _load_validated_ledger(path: Path, expected_workflow: str) -> dict[str, Any]:
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RunVerificationError(f"cannot read ledger {path}") from exc
    ledger = _as_mapping(raw, "ledger")
    if ledger.get("schema_version") != LEDGER_SCHEMA_VERSION:
        raise RunVerificationError("ledger schema version is invalid")

    revision = ledger.get("revision")
    if not isinstance(revision, str):
        raise RunVerificationError("ledger revision is invalid")
    _require_revision(revision)
    if ledger.get("workflow") != expected_workflow:
        raise RunVerificationError(
            f"ledger workflow {ledger.get('workflow')!r} does not match {expected_workflow!r}"
        )

    run = _as_mapping(ledger.get("run"), "ledger run")
    if run.get("head_sha") != revision:
        raise RunVerificationError("ledger run SHA does not match ledger revision")
    if run.get("status") != "completed" or run.get("conclusion") != "success":
        raise RunVerificationError("ledger run is not completed successfully")
    database_id = run.get("database_id")
    if not isinstance(database_id, int) or isinstance(database_id, bool):
        raise RunVerificationError("ledger run ID is invalid")

    ledger_jobs = _as_sequence(ledger.get("jobs"), "ledger jobs")
    reconstructed_jobs = [
        {
            "databaseId": _as_mapping(job, "ledger job").get("database_id"),
            "name": _as_mapping(job, "ledger job").get("name"),
            "status": _as_mapping(job, "ledger job").get("status"),
            "conclusion": _as_mapping(job, "ledger job").get("conclusion"),
        }
        for job in ledger_jobs
    ]
    _validate_required_jobs(reconstructed_jobs)
    return ledger


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command", required=True)

    collect_parser = subparsers.add_parser("collect")
    collect_parser.add_argument("--revision", required=True)
    collect_parser.add_argument("--workflow", required=True)
    collect_parser.add_argument("--output", required=True, type=Path)

    print_run_parser = subparsers.add_parser("print-run-id")
    print_run_parser.add_argument("--ledger", required=True, type=Path)
    print_run_parser.add_argument("--workflow", required=True)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    arguments = _parser().parse_args(argv)
    if arguments.command == "collect":
        collect(
            revision=arguments.revision,
            workflow=arguments.workflow,
            output=arguments.output,
        )
        return 0
    if arguments.command == "print-run-id":
        ledger = _load_validated_ledger(arguments.ledger, arguments.workflow)
        print(ledger["run"]["database_id"])
        return 0
    raise AssertionError(f"unhandled command: {arguments.command}")


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except RunVerificationError as exc:
        print(f"native CI verification failed: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc
