"""Policy-neutral planning for ordered multi-sheet operations."""

from __future__ import annotations

from collections.abc import Callable, Iterable
from dataclasses import dataclass
from enum import StrEnum
from typing import Protocol

from messy_xlsx.models import SheetInfo
from messy_xlsx.parsing.parse_plan import ParsePlan

_MAX_FAILURE_GRAPH_NODES = 10_000


class SheetSelectionOptions(Protocol):
    """Minimal legacy option surface consumed by the shared planner."""

    sheets: list[str] | None
    sheet_filter: Callable[[SheetInfo], bool] | None


class PlannedSheetState(StrEnum):
    """Unambiguous outcome of one sheet's planning pass."""

    READY = "ready"
    SKIPPED = "skipped"
    ERROR = "error"


class PlanningFailureStage(StrEnum):
    """Stage at which an otherwise ordinary planning failure occurred."""

    ANALYSIS = "analysis"
    COMPILE = "compile"


@dataclass(frozen=True, slots=True)
class PlannedSheet:
    """One immutable planning decision in workbook order."""

    name: str
    info: SheetInfo
    state: PlannedSheetState
    parse_plan: ParsePlan | None = None
    error: BaseException | None = None
    failure_stage: PlanningFailureStage | None = None

    def __post_init__(self) -> None:
        ready = self.state is PlannedSheetState.READY
        failed = self.state is PlannedSheetState.ERROR
        if ready != (self.parse_plan is not None):
            raise ValueError("ready sheets require exactly one parse plan")
        if failed != (self.error is not None):
            raise ValueError("failed planning requires exactly one error")
        if failed != (self.failure_stage is not None):
            raise ValueError("failed planning requires exactly one failure stage")
        if self.parse_plan is not None and self.error is not None:
            raise ValueError("planned sheets cannot contain both a plan and an error")


class SheetPlanner:
    """Analyze, select, and compile sheets once without owning adapter policy."""

    def __init__(
        self,
        analyze: Callable[[str], SheetInfo],
        compile_selected: Callable[[str, SheetInfo], ParsePlan],
        *,
        should_propagate: Callable[[BaseException], bool],
        analysis_failure_info: Callable[[str, Exception], SheetInfo],
    ) -> None:
        self._analyze = analyze
        self._compile_selected = compile_selected
        self._should_propagate = should_propagate
        self._analysis_failure_info = analysis_failure_info

    def plan(
        self,
        names: Iterable[str],
        *,
        options: SheetSelectionOptions | None = None,
        select_all: bool = False,
        compile_outputs: bool = True,
    ) -> tuple[PlannedSheet, ...]:
        """Return one explicit planning outcome for every workbook sheet."""
        planned: list[PlannedSheet] = []
        explicit_names = _freeze_explicit_names(options)
        for name in names:
            try:
                info = self._analyze(name)
            except BaseException as error:
                if self._must_propagate(error):
                    raise
                assert isinstance(error, Exception)
                planned.append(
                    PlannedSheet(
                        name=name,
                        info=self._analysis_failure_info(name, error),
                        state=PlannedSheetState.ERROR,
                        error=_clear_exception_tracebacks(error),
                        failure_stage=PlanningFailureStage.ANALYSIS,
                    )
                )
                continue

            selected = select_all or self._selected(info, options, explicit_names)
            if not selected or not compile_outputs:
                planned.append(
                    PlannedSheet(
                        name=name,
                        info=info,
                        state=PlannedSheetState.SKIPPED,
                    )
                )
                continue

            # ``SheetInfo`` is intentionally mutable.  The legacy adapter uses
            # the possibly filter-mutated name for both parsing and result
            # keys, so the immutable plan must capture that effective target.
            effective_name = info.name
            try:
                parse_plan = self._compile_selected(effective_name, info)
            except BaseException as error:
                if self._must_propagate(error):
                    raise
                assert isinstance(error, Exception)
                planned.append(
                    PlannedSheet(
                        name=effective_name,
                        info=info,
                        state=PlannedSheetState.ERROR,
                        error=_clear_exception_tracebacks(error),
                        failure_stage=PlanningFailureStage.COMPILE,
                    )
                )
                continue
            planned.append(
                PlannedSheet(
                    name=effective_name,
                    info=info,
                    state=PlannedSheetState.READY,
                    parse_plan=parse_plan,
                )
            )
        return tuple(planned)

    def _must_propagate(self, error: BaseException) -> bool:
        return not isinstance(error, Exception) or self._should_propagate(error)

    @staticmethod
    def _selected(
        info: SheetInfo,
        options: SheetSelectionOptions | None,
        explicit_names: frozenset[str] | None,
    ) -> bool:
        if options is None:
            return True
        # This exact order is a frozen legacy contract.
        if info.skip_reason:
            return False
        if options.sheet_filter is not None and not options.sheet_filter(info):
            return False
        if explicit_names is not None:
            return not explicit_names or info.name in explicit_names
        return not options.sheets or info.name in options.sheets


def _freeze_explicit_names(
    options: SheetSelectionOptions | None,
) -> frozenset[str] | None:
    """Snapshot ordinary selection lists when no filter can mutate them."""
    if options is None or options.sheet_filter is not None:
        return None
    sheets = options.sheets
    if type(sheets) is not list:
        return None
    normalized: list[str] = []
    for name in sheets:
        if not isinstance(name, str):
            return None
        normalized.append(str.__str__(name))
    return frozenset(normalized)


def _clear_exception_tracebacks(error: BaseException) -> BaseException:
    """Drop traceback frames from a bounded failure graph before retention."""
    stack = [error]
    seen: set[int] = set()
    while stack and len(seen) < _MAX_FAILURE_GRAPH_NODES:
        candidate = stack.pop()
        identity = id(candidate)
        if identity in seen:
            continue
        seen.add(identity)
        try:
            cause = BaseException.__getattribute__(candidate, "__cause__")
            context = BaseException.__getattribute__(candidate, "__context__")
            nested = (
                BaseException.__getattribute__(candidate, "exceptions")
                if isinstance(candidate, BaseExceptionGroup)
                else ()
            )
        except BaseException:
            cause = context = None
            nested = ()
        if isinstance(cause, BaseException):
            stack.append(cause)
        if isinstance(context, BaseException):
            stack.append(context)
        if isinstance(nested, tuple):
            stack.extend(item for item in nested if isinstance(item, BaseException))
        try:
            BaseException.__setattr__(candidate, "__traceback__", None)
        except BaseException:
            pass
    return error


__all__ = [
    "PlannedSheet",
    "PlannedSheetState",
    "PlanningFailureStage",
    "SheetPlanner",
    "SheetSelectionOptions",
]
