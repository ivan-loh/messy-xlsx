from __future__ import annotations

import importlib.util
import json
import sys
from collections.abc import Sequence
from pathlib import Path
from types import ModuleType
from typing import Any

import pytest
import yaml

PROJECT_ROOT = Path(__file__).resolve().parents[2]
SCRIPT_PATH = PROJECT_ROOT / "scripts" / "verify_native_ci.py"
WORKFLOW_PATH = PROJECT_ROOT / ".github" / "workflows" / "native-abi.yml"
REVISION = "a" * 40
PLATFORMS = (
    "manylinux-x86_64",
    "manylinux-aarch64",
    "musllinux-x86_64",
    "musllinux-aarch64",
    "macos-x86_64",
    "macos-arm64",
    "windows-amd64",
)
PYTHONS = ("cp311", "cp312", "cp313", "cp314")
REQUIRED_JOB_NAMES = {
    *(f"native-wheel / {platform}" for platform in PLATFORMS),
    *(f"abi3-smoke / {platform} / {python}" for platform in PLATFORMS for python in PYTHONS),
}


def _load_verifier() -> ModuleType:
    assert SCRIPT_PATH.is_file(), "Task 2 must provide scripts/verify_native_ci.py"
    spec = importlib.util.spec_from_file_location("verify_native_ci_under_test", SCRIPT_PATH)
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


class RecordedGH:
    def __init__(
        self,
        run_lists: Sequence[list[dict[str, Any]]],
        run_view: dict[str, Any] | None = None,
    ) -> None:
        self._run_lists = list(run_lists)
        self._run_view = run_view
        self.calls: list[tuple[str, ...]] = []

    def __call__(self, arguments: Sequence[str]) -> object:
        call = tuple(arguments)
        self.calls.append(call)
        if call[:2] == ("run", "list"):
            if len(self._run_lists) > 1:
                return self._run_lists.pop(0)
            return self._run_lists[0]
        if call[:2] == ("run", "view") and self._run_view is not None:
            return self._run_view
        raise AssertionError(f"unexpected gh call: {call!r}")


class RecordedClock:
    def __init__(self) -> None:
        self.now = 0.0
        self.sleeps: list[float] = []

    def monotonic(self) -> float:
        return self.now

    def sleep(self, seconds: float) -> None:
        self.sleeps.append(seconds)
        self.now += seconds


def _run(
    *,
    head_sha: str = REVISION,
    status: str = "completed",
    conclusion: str | None = "success",
    database_id: int = 4201,
) -> dict[str, Any]:
    return {
        "databaseId": database_id,
        "headSha": head_sha,
        "status": status,
        "conclusion": conclusion,
        "url": f"https://example.invalid/runs/{database_id}",
    }


def _successful_jobs() -> list[dict[str, Any]]:
    return [
        {
            "name": name,
            "status": "completed",
            "conclusion": "success",
            "databaseId": index,
        }
        for index, name in enumerate(sorted(REQUIRED_JOB_NAMES), start=1)
    ]


def _view(
    *,
    head_sha: str = REVISION,
    status: str = "completed",
    conclusion: str | None = "success",
    jobs: list[dict[str, Any]] | None = None,
) -> dict[str, Any]:
    return {
        **_run(head_sha=head_sha, status=status, conclusion=conclusion),
        "jobs": _successful_jobs() if jobs is None else jobs,
    }


def test_collect_waits_for_visibility_and_completion_without_real_sleep(
    tmp_path: Path,
) -> None:
    verifier = _load_verifier()
    output = tmp_path / "ledger.json"
    gh = RecordedGH(
        [
            [],
            [_run(status="queued", conclusion=None)],
            [_run()],
        ],
        _view(),
    )
    clock = RecordedClock()

    ledger = verifier.collect(
        revision=REVISION,
        workflow="native-abi.yml",
        output=output,
        gh_json=gh,
        monotonic=clock.monotonic,
        sleep=clock.sleep,
    )

    assert clock.sleeps == [30.0, 30.0]
    assert ledger["revision"] == REVISION
    assert ledger["workflow"] == "native-abi.yml"
    assert ledger["run"]["database_id"] == 4201
    assert {job["name"] for job in ledger["jobs"]} == REQUIRED_JOB_NAMES
    assert json.loads(output.read_text(encoding="utf-8")) == ledger


@pytest.mark.parametrize("conclusion", ["failure", "cancelled", "skipped", "neutral"])
def test_collect_rejects_terminal_nonsuccess_conclusion(
    tmp_path: Path,
    conclusion: str,
) -> None:
    verifier = _load_verifier()
    gh = RecordedGH([[_run(conclusion=conclusion)]])

    with pytest.raises(verifier.RunVerificationError, match=conclusion):
        verifier.collect(
            revision=REVISION,
            workflow="native-abi.yml",
            output=tmp_path / "ledger.json",
            gh_json=gh,
        )


def test_collect_times_out_on_nonterminal_run_without_sleeping(
    tmp_path: Path,
) -> None:
    verifier = _load_verifier()
    gh = RecordedGH([[_run(status="in_progress", conclusion=None)]])
    clock = RecordedClock()

    with pytest.raises(verifier.RunVerificationError, match="timed out"):
        verifier.collect(
            revision=REVISION,
            workflow="native-abi.yml",
            output=tmp_path / "ledger.json",
            gh_json=gh,
            monotonic=clock.monotonic,
            sleep=clock.sleep,
            timeout_seconds=60,
        )

    assert clock.sleeps == [30.0, 30.0]


def test_collect_rejects_wrong_sha_run(
    tmp_path: Path,
) -> None:
    verifier = _load_verifier()
    gh = RecordedGH([[_run(head_sha="b" * 40)]])
    clock = RecordedClock()

    with pytest.raises(verifier.RunVerificationError, match="timed out"):
        verifier.collect(
            revision=REVISION,
            workflow="native-abi.yml",
            output=tmp_path / "ledger.json",
            gh_json=gh,
            monotonic=clock.monotonic,
            sleep=clock.sleep,
            timeout_seconds=30,
        )


def test_collect_rechecks_sha_from_run_view(
    tmp_path: Path,
) -> None:
    verifier = _load_verifier()
    gh = RecordedGH([[_run()]], _view(head_sha="b" * 40))

    with pytest.raises(verifier.RunVerificationError, match="SHA"):
        verifier.collect(
            revision=REVISION,
            workflow="native-abi.yml",
            output=tmp_path / "ledger.json",
            gh_json=gh,
        )


def test_collect_rejects_incomplete_required_job_matrix(
    tmp_path: Path,
) -> None:
    verifier = _load_verifier()
    jobs = _successful_jobs()
    jobs.pop()
    gh = RecordedGH([[_run()]], _view(jobs=jobs))

    with pytest.raises(verifier.RunVerificationError, match="missing required jobs"):
        verifier.collect(
            revision=REVISION,
            workflow="native-abi.yml",
            output=tmp_path / "ledger.json",
            gh_json=gh,
        )


def test_collect_rejects_unsuccessful_required_job(
    tmp_path: Path,
) -> None:
    verifier = _load_verifier()
    jobs = _successful_jobs()
    jobs[0] = {**jobs[0], "conclusion": "failure"}
    gh = RecordedGH([[_run()]], _view(jobs=jobs))

    with pytest.raises(verifier.RunVerificationError, match="required job"):
        verifier.collect(
            revision=REVISION,
            workflow="native-abi.yml",
            output=tmp_path / "ledger.json",
            gh_json=gh,
        )


def test_print_run_id_validates_ledger_and_workflow(
    tmp_path: Path,
    capsys: pytest.CaptureFixture[str],
) -> None:
    verifier = _load_verifier()
    output = tmp_path / "ledger.json"
    verifier.collect(
        revision=REVISION,
        workflow="native-abi.yml",
        output=output,
        gh_json=RecordedGH([[_run()]], _view()),
    )

    assert (
        verifier.main(
            [
                "print-run-id",
                "--ledger",
                str(output),
                "--workflow",
                "native-abi.yml",
            ]
        )
        == 0
    )
    assert capsys.readouterr().out == "4201\n"

    with pytest.raises(verifier.RunVerificationError, match="workflow"):
        verifier.main(
            [
                "print-run-id",
                "--ledger",
                str(output),
                "--workflow",
                "native-artifacts.yml",
            ]
        )


def _load_workflow() -> dict[str, Any]:
    assert WORKFLOW_PATH.is_file(), "Task 2 must provide native-abi.yml"
    loaded = yaml.load(WORKFLOW_PATH.read_text(encoding="utf-8"), Loader=yaml.BaseLoader)
    assert isinstance(loaded, dict)
    return loaded


def test_native_abi_workflow_is_dispatchable_and_reusable() -> None:
    workflow = _load_workflow()

    triggers = workflow["on"]
    assert triggers["push"]["branches"] == ["**"]
    assert "workflow_dispatch" in triggers
    assert "workflow_call" in triggers


def test_native_abi_workflow_covers_every_claimed_wheel_and_runtime() -> None:
    workflow = _load_workflow()

    build_matrix = workflow["jobs"]["build-native"]["strategy"]["matrix"]["include"]
    assert {entry["platform"] for entry in build_matrix} == set(PLATFORMS)

    smoke_matrix = workflow["jobs"]["abi3-smoke"]["strategy"]["matrix"]["include"]
    assert {(entry["platform"], entry["python"]) for entry in smoke_matrix} == {
        (platform, python) for platform in PLATFORMS for python in PYTHONS
    }
