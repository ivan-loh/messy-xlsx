"""Handler registry for routing files to appropriate format handlers."""

# ============================================================================
# Imports
# ============================================================================

from collections.abc import Iterable, Iterator
from contextlib import contextmanager
from enum import Enum
from types import (
    BuiltinFunctionType,
    FunctionType,
    MemberDescriptorType,
    MethodDescriptorType,
    ModuleType,
    WrapperDescriptorType,
)
from typing import cast

import pandas as pd

from messy_xlsx._fallback_signals import _is_fallback_blocked
from messy_xlsx._source import SourceHandle, SourceInput
from messy_xlsx.detection.format_detector import FormatDetector
from messy_xlsx.exceptions import FormatError
from messy_xlsx.models import FormatInfo
from messy_xlsx.parsing.base_handler import FormatHandler, ParseOptions
from messy_xlsx.parsing.csv_handler import CSVHandler
from messy_xlsx.parsing.xls_handler import XLSHandler
from messy_xlsx.parsing.xlsx_handler import XLSXHandler

_BUILTIN_HANDLER_TYPES = (XLSXHandler, XLSHandler, CSVHandler)
_MAX_FINGERPRINT_NODES = 10_000
_MAX_FINGERPRINT_COST = 10_000
_REGISTRY_STATE_NAMES = frozenset(
    {
        "handlers",
        "detector",
        "_builtin_handler_instances",
        "_builtin_detector_instance",
        "_builtin_component_token",
    }
)
_CANONICAL_BEHAVIOR_TYPES: tuple[type[object], ...] = ()
_CANONICAL_CLASS_BEHAVIOR: object | None = None
_CANONICAL_COMPONENT_SHAPE: object | None = None
_FINGERPRINT_BOOTSTRAP_GLOBALS = frozenset(
    {
        "_CANONICAL_BEHAVIOR_TYPES",
        "_CANONICAL_CLASS_BEHAVIOR",
        "_CANONICAL_COMPONENT_SHAPE",
    }
)


class _FingerprintError(RuntimeError):
    """Raised when an extension graph cannot be inspected safely."""


class _FingerprintBudget:
    def __init__(self) -> None:
        self._cost = 0

    def _charge(self, units: int = 1) -> None:
        self._cost += units
        if self._cost > _MAX_FINGERPRINT_COST:
            raise _FingerprintError("fingerprint budget exceeded")


class _CompositionFingerprinter(_FingerprintBudget):
    """Build a bounded recursive identity-and-state token for owned components."""

    def __init__(self, *, include_identity: bool) -> None:
        super().__init__()
        self._seen: dict[int, int] = {}
        self._include_identity = include_identity

    def token(self, value: object) -> object:  # noqa: C901
        self._charge()
        if isinstance(value, bytes):
            self._charge(len(value))
            return (bytes, value)
        if value is None or isinstance(value, bool):
            return (type(value), value)
        if isinstance(value, int):
            self._charge(max(1, (value.bit_length() + 7) // 8))
            return (int, value)
        if isinstance(value, str):
            self._charge(len(value))
            return (str, value)
        if isinstance(value, float):
            return (float, value.hex())
        if isinstance(value, complex):
            return (complex, value.real.hex(), value.imag.hex())
        if isinstance(value, Enum):
            return ("enum", type(value), value.name)
        if isinstance(value, type):
            return ("type", value)
        if isinstance(
            value,
            (
                FunctionType,
                BuiltinFunctionType,
                MethodDescriptorType,
                WrapperDescriptorType,
            ),
        ):
            return ("callable", type(value), id(value))

        identity = id(value)
        if identity in self._seen:
            return ("ref", self._seen[identity])
        if len(self._seen) >= _MAX_FINGERPRINT_NODES:
            raise _FingerprintError("fingerprint budget exceeded")
        node = len(self._seen)
        self._seen[identity] = node

        if type(value) is dict:
            self._charge(len(value) * 2)
            return (
                "dict",
                node,
                identity if self._include_identity else None,
                frozenset((self.token(key), self.token(item)) for key, item in value.items()),
            )
        if type(value) is list:
            self._charge(len(value))
            return (
                "list",
                node,
                identity if self._include_identity else None,
                tuple(self.token(item) for item in value),
            )
        if type(value) is tuple:
            self._charge(len(value))
            return (
                "tuple",
                node,
                identity if self._include_identity else None,
                tuple(self.token(item) for item in value),
            )
        if type(value) is set:
            self._charge(len(value))
            return (
                "set",
                node,
                identity if self._include_identity else None,
                frozenset(self.token(item) for item in value),
            )
        if type(value) is frozenset:
            self._charge(len(value))
            return (
                "frozenset",
                node,
                identity if self._include_identity else None,
                frozenset(self.token(item) for item in value),
            )
        if type(value) is bytearray:
            self._charge(len(value))
            return (
                "bytearray",
                node,
                identity if self._include_identity else None,
                bytes(value),
            )
        if type(value) is memoryview:
            self._charge(value.nbytes)
            return (
                "memoryview",
                node,
                identity if self._include_identity else None,
                value.tobytes(),
            )

        state = dict(vars(value)) if hasattr(value, "__dict__") else {}
        for owner in type(value).__mro__:
            for name, descriptor in vars(owner).items():
                if isinstance(descriptor, MemberDescriptorType):
                    try:
                        state.setdefault(name, descriptor.__get__(value, type(value)))
                    except AttributeError:
                        continue
        self._charge(len(state))
        return (
            "object",
            node,
            identity if self._include_identity else None,
            type(value),
            frozenset((name, self.token(item)) for name, item in state.items()),
        )


class _BehaviorFingerprinter(_FingerprintBudget):
    """Snapshot raw class attributes without binding or invoking descriptors."""

    def __init__(self) -> None:
        super().__init__()
        self._seen: dict[int, int] = {}

    def token(self, value: object) -> object:  # noqa: C901
        self._charge()
        if isinstance(value, bytes):
            self._charge(len(value))
            return (bytes, value)
        if value is None or isinstance(value, bool):
            return (type(value), value)
        if isinstance(value, int):
            self._charge(max(1, (value.bit_length() + 7) // 8))
            return (int, value)
        if isinstance(value, str):
            self._charge(len(value))
            return (str, value)
        if isinstance(value, float):
            return (float, value.hex())
        if isinstance(value, complex):
            return (complex, value.real.hex(), value.imag.hex())
        if isinstance(value, Enum):
            return ("enum", type(value), value.name)
        if isinstance(value, FunctionType):
            identity = id(value)
            if identity in self._seen:
                return ("ref", self._seen[identity])
            if len(self._seen) >= _MAX_FINGERPRINT_NODES:
                raise _FingerprintError("fingerprint budget exceeded")
            node = len(self._seen)
            self._seen[identity] = node
            closure: list[tuple[str, object]] = []
            self._charge(len(value.__closure__ or ()) + len(value.__code__.co_names) + 4)
            for name, cell in zip(
                value.__code__.co_freevars,
                value.__closure__ or (),
                strict=True,
            ):
                self._charge(len(name))
                try:
                    contents = cell.cell_contents
                except ValueError:
                    closure.append((name, ("empty-cell",)))
                else:
                    closure.append((name, self.token(contents)))
            referenced_globals = self._referenced_globals(value)
            return (
                "function",
                node,
                id(value),
                id(value.__code__),
                tuple(closure),
                referenced_globals,
                self.token(value.__defaults__),
                self.token(value.__kwdefaults__),
                self.token(value.__annotations__),
                self.token(vars(value)),
            )
        if isinstance(value, staticmethod):
            self._charge()
            return ("staticmethod", self.token(value.__func__))
        if isinstance(value, classmethod):
            self._charge()
            return ("classmethod", self.token(value.__func__))
        if isinstance(value, property):
            self._charge(3)
            return (
                "property",
                self.token(value.fget),
                self.token(value.fset),
                self.token(value.fdel),
            )
        if isinstance(value, type):
            return ("type", value, id(value))
        if isinstance(value, ModuleType):
            return ("module", value.__name__, id(value))
        if isinstance(
            value,
            (
                BuiltinFunctionType,
                MethodDescriptorType,
                WrapperDescriptorType,
            ),
        ):
            return ("callable", type(value), id(value))
        if hasattr(type(value), "__get__"):
            return ("descriptor", type(value), id(value))

        identity = id(value)
        if identity in self._seen:
            return ("ref", self._seen[identity])
        if len(self._seen) >= _MAX_FINGERPRINT_NODES:
            raise _FingerprintError("fingerprint budget exceeded")
        node = len(self._seen)
        self._seen[identity] = node

        if type(value) is dict:
            self._charge(len(value) * 2)
            return (
                "dict",
                node,
                frozenset((self.token(key), self.token(item)) for key, item in value.items()),
            )
        if type(value) is list:
            self._charge(len(value))
            return ("list", node, tuple(self.token(item) for item in value))
        if type(value) is tuple:
            self._charge(len(value))
            return ("tuple", node, tuple(self.token(item) for item in value))
        if type(value) is set:
            self._charge(len(value))
            return ("set", node, frozenset(self.token(item) for item in value))
        if type(value) is frozenset:
            self._charge(len(value))
            return ("frozenset", node, frozenset(self.token(item) for item in value))
        if type(value) is bytearray:
            self._charge(len(value))
            return ("bytearray", node, bytes(value))
        if type(value) is memoryview:
            self._charge(value.nbytes)
            return ("memoryview", node, value.tobytes())
        return ("opaque", type(value), identity)

    def _referenced_globals(
        self,
        function: FunctionType,
    ) -> tuple[tuple[str, object], ...]:
        """Fingerprint only globals named by this function's bytecode."""
        references: list[tuple[str, object]] = []
        namespace = function.__globals__
        for name in sorted(set(function.__code__.co_names)):
            if name in _FINGERPRINT_BOOTSTRAP_GLOBALS or name not in namespace:
                continue
            self._charge(len(name))
            references.append(
                (name, self._global_reference_token(dict.__getitem__(namespace, name)))
            )
        return tuple(references)

    def _global_reference_token(self, value: object) -> object:
        """Recurse only into project functions and bounded data globals."""
        if isinstance(value, FunctionType):
            if isinstance(value.__module__, str) and value.__module__.startswith("messy_xlsx"):
                return self.token(value)
            return ("external-function", id(value))
        if isinstance(value, type):
            try:
                module = type.__getattribute__(value, "__module__")
            except BaseException:
                module = "<unknown>"
            return ("project-type" if module.startswith("messy_xlsx") else "type", id(value))
        if isinstance(value, ModuleType):
            return ("module", value.__name__, id(value))
        if isinstance(
            value,
            (
                BuiltinFunctionType,
                MethodDescriptorType,
                WrapperDescriptorType,
            ),
        ):
            return ("callable", type(value), id(value))
        if (
            value is None
            or isinstance(
                value,
                (bool, int, float, complex, str, bytes, Enum),
            )
            or type(value) in {dict, list, tuple, set, frozenset, bytearray, memoryview}
        ):
            return self.token(value)
        return ("global", type(value), id(value))


def _component_token(
    handlers: list[FormatHandler],
    detector: FormatDetector,
    *,
    include_identity: bool = True,
) -> object:
    fingerprinter = _CompositionFingerprinter(include_identity=include_identity)
    return (
        "components",
        fingerprinter.token(handlers),
        fingerprinter.token(detector),
    )


def _project_component_types(  # noqa: C901
    components: tuple[object, ...],
) -> tuple[type[object], ...]:
    """Discover project-owned instance types without following class descriptors."""
    seen: set[int] = set()
    discovered: set[type[object]] = {HandlerRegistry}

    def visit(value: object) -> None:  # noqa: C901
        if value is None or isinstance(
            value,
            (bool, int, float, complex, str, bytes, Enum, type, FunctionType),
        ):
            return
        identity = id(value)
        if identity in seen or len(seen) >= _MAX_FINGERPRINT_NODES:
            return
        seen.add(identity)
        if type(value) is dict:
            for key, item in value.items():
                visit(key)
                visit(item)
            return
        if type(value) in {list, tuple, set, frozenset}:
            for item in cast(Iterable[object], value):
                visit(item)
            return
        if type(value) in {bytearray, memoryview}:
            return
        value_type = type(value)
        for owner in value_type.__mro__:
            if owner.__module__.startswith("messy_xlsx"):
                discovered.add(owner)
        if hasattr(value, "__dict__"):
            for item in vars(value).values():
                visit(item)

    for component in components:
        visit(component)
    return tuple(sorted(discovered, key=lambda item: (item.__module__, item.__qualname__)))


def _class_behavior_token(component_types: tuple[type[object], ...]) -> object:
    """Fingerprint every raw attribute on every built-in project class/MRO."""
    fingerprinter = _BehaviorFingerprinter()
    return tuple(
        (
            component_type,
            tuple(
                (name, fingerprinter.token(value))
                for name, value in sorted(vars(component_type).items())
            ),
        )
        for component_type in component_types
    )


# ============================================================================
# Core
# ============================================================================


class HandlerRegistry:
    """Registry of format handlers with automatic detection and fallback."""

    _accepts_source_handle = True

    def __init__(
        self,
        handlers: list[FormatHandler] | None = None,
        detector: FormatDetector | None = None,
    ) -> None:
        """Initialize a registry, optionally with caller-supplied components."""
        uses_default_components = handlers is None and detector is None
        self.handlers = (
            list(handlers)
            if handlers is not None
            else [
                XLSXHandler(),
                XLSHandler(),
                CSVHandler(),
            ]
        )
        self.detector = detector or FormatDetector()
        self._builtin_handler_instances: tuple[FormatHandler, ...] | None = None
        self._builtin_detector_instance: FormatDetector | None = None
        self._builtin_component_token: object | None = None
        if uses_default_components and type(self) is HandlerRegistry:
            try:
                component_shape = _component_token(
                    self.handlers,
                    self.detector,
                    include_identity=False,
                )
                if (
                    _CANONICAL_COMPONENT_SHAPE is None
                    or component_shape == _CANONICAL_COMPONENT_SHAPE
                ):
                    self._builtin_handler_instances = tuple(self.handlers)
                    self._builtin_detector_instance = self.detector
                    self._builtin_component_token = _component_token(
                        self.handlers,
                        self.detector,
                    )
            except Exception:
                pass

    def _uses_builtin_components(self) -> bool:
        """Return whether no caller extension can be bypassed by new backends."""
        if (
            self._builtin_handler_instances is None
            or self._builtin_component_token is None
            or _CANONICAL_CLASS_BEHAVIOR is None
        ):
            return False
        try:
            return (
                type(self) is HandlerRegistry
                and set(vars(self)) == _REGISTRY_STATE_NAMES
                and tuple(type(handler) for handler in self.handlers) == _BUILTIN_HANDLER_TYPES
                and len(self.handlers) == len(self._builtin_handler_instances)
                and all(
                    current is original
                    for current, original in zip(
                        self.handlers,
                        self._builtin_handler_instances,
                        strict=True,
                    )
                )
                and type(self.detector) is FormatDetector
                and self.detector is self._builtin_detector_instance
                and _component_token(self.handlers, self.detector) == self._builtin_component_token
                and _class_behavior_token(_CANONICAL_BEHAVIOR_TYPES) == _CANONICAL_CLASS_BEHAVIOR
            )
        except Exception:
            return False

    def register_handler(self, handler: FormatHandler, priority: int = -1) -> None:
        """Register a custom handler."""
        if priority < 0:
            self.handlers.append(handler)
        else:
            self.handlers.insert(priority, handler)

    def get_handler(self, format_type: str) -> FormatHandler | None:
        """Get handler for a specific format type."""
        for handler in self.handlers:
            if handler.can_handle(format_type):
                return handler
        return None

    def detect_format(
        self,
        file_source: SourceInput | SourceHandle,
        filename: str | None = None,
    ) -> FormatInfo:
        """Detect file format."""
        with self._source_handle(file_source, filename) as source:
            return self._detect(source, filename)

    def parse(  # noqa: C901
        self,
        file_source: SourceInput | SourceHandle,
        sheet: str | None = None,
        options: ParseOptions | None = None,
        format_type: str | None = None,
    ) -> pd.DataFrame:
        """Parse file with automatic format detection and fallback."""
        with self._source_handle(file_source) as source:
            file_desc = source.description
            options = options or ParseOptions()

            if format_type is None:
                format_info = self._detect(source)
                format_type = format_info.format_type

            handler = self.get_handler(format_type)

            if handler is None:
                raise FormatError(
                    f"No handler available for format: {format_type}",
                    file_path=file_desc,
                    detected_format=format_type,
                )

            try:
                return self._parse_with(handler, source, sheet, options)
            except (PermissionError, FileNotFoundError, MemoryError):
                raise
            except Exception as error:
                if _is_fallback_blocked(error):
                    raise
                pass

            attempted_handlers = [handler.__class__.__name__]
            for fallback_handler in self.handlers:
                if fallback_handler == handler or not fallback_handler.can_handle(format_type):
                    continue

                attempted_handlers.append(fallback_handler.__class__.__name__)
                try:
                    return self._parse_with(fallback_handler, source, sheet, options)
                except (PermissionError, FileNotFoundError, MemoryError):
                    raise
                except Exception as error:
                    if _is_fallback_blocked(error):
                        raise
                    continue

            name = source.path.name if source.path is not None else file_desc
            raise FormatError(
                f"All handlers failed for {name}",
                file_path=file_desc,
                detected_format=format_type,
                attempted_formats=attempted_handlers,
            )

    def get_sheet_names(  # noqa: C901
        self,
        file_source: SourceInput | SourceHandle,
        format_type: str | None = None,
    ) -> list[str]:
        """Get sheet names from file."""
        with self._source_handle(file_source) as source:
            if format_type is None:
                format_info = self._detect(source)
                format_type = format_info.format_type

            handler = self.get_handler(format_type)

            if handler is None:
                return ["Sheet1"]

            try:
                return self._sheet_names_with(handler, source)
            except (PermissionError, FileNotFoundError, MemoryError):
                raise
            except Exception as error:
                if _is_fallback_blocked(error):
                    raise
                pass

            for fallback_handler in self.handlers:
                if fallback_handler == handler or not fallback_handler.can_handle(format_type):
                    continue

                try:
                    return self._sheet_names_with(fallback_handler, source)
                except (PermissionError, FileNotFoundError, MemoryError):
                    raise
                except Exception as error:
                    if _is_fallback_blocked(error):
                        raise
                    continue

            return ["Sheet1"]

    def validate(
        self,
        file_source: SourceInput | SourceHandle,
        format_type: str | None = None,
    ) -> tuple[bool, str | None]:
        """Validate that file can be parsed."""
        with self._source_handle(file_source) as source:
            if format_type is None:
                try:
                    format_info = self._detect(source)
                    format_type = format_info.format_type
                except FormatError as e:
                    return False, str(e)

            if format_type == "unknown":
                return False, "Unknown file format"

            handler = self.get_handler(format_type)

            if handler is None:
                return False, f"No handler for format: {format_type}"

            return self._validate_with(handler, source)

    @contextmanager
    def _source_handle(
        self,
        source: SourceInput | SourceHandle,
        filename: str | None = None,
    ) -> Iterator[SourceHandle]:
        """Keep caller-owned handles alive and close local snapshots."""
        created = not isinstance(source, SourceHandle)
        handle = SourceHandle.coerce(source, filename=filename)
        try:
            yield handle
        finally:
            if created:
                handle.close()

    def _detect(self, source: SourceHandle, filename: str | None = None) -> FormatInfo:
        accepts_handle = bool(type(self.detector).__dict__.get("_accepts_source_handle", False))
        if accepts_handle:
            if filename is None:
                return self.detector.detect(source)
            return self.detector.detect(source, filename=filename)
        with source.open_legacy() as legacy_source:
            if filename is None:
                return self.detector.detect(legacy_source)
            return self.detector.detect(legacy_source, filename=filename)

    @staticmethod
    def _handler_accepts_source_handle(handler: FormatHandler) -> bool:
        return bool(type(handler).__dict__.get("_accepts_source_handle", False))

    def _parse_with(
        self,
        handler: FormatHandler,
        source: SourceHandle,
        sheet: str | None,
        options: ParseOptions,
    ) -> pd.DataFrame:
        if self._handler_accepts_source_handle(handler):
            return handler.parse(source, sheet, options)  # type: ignore[arg-type]
        with source.open_legacy() as legacy_source:
            return handler.parse(legacy_source, sheet, options)

    def _sheet_names_with(
        self,
        handler: FormatHandler,
        source: SourceHandle,
    ) -> list[str]:
        if self._handler_accepts_source_handle(handler):
            return handler.get_sheet_names(source)  # type: ignore[arg-type]
        with source.open_legacy() as legacy_source:
            return handler.get_sheet_names(legacy_source)

    def _validate_with(
        self,
        handler: FormatHandler,
        source: SourceHandle,
    ) -> tuple[bool, str | None]:
        if self._handler_accepts_source_handle(handler):
            return handler.validate(source)  # type: ignore[arg-type]
        with source.open_legacy() as legacy_source:
            return handler.validate(legacy_source)


# ============================================================================
# Module Entrypoint
# ============================================================================

_behavior_probe = HandlerRegistry()
_CANONICAL_BEHAVIOR_TYPES = _project_component_types(
    (HandlerRegistry, _behavior_probe.handlers, _behavior_probe.detector)
)
_CANONICAL_CLASS_BEHAVIOR = _class_behavior_token(_CANONICAL_BEHAVIOR_TYPES)
_CANONICAL_COMPONENT_SHAPE = _component_token(
    _behavior_probe.handlers,
    _behavior_probe.detector,
    include_identity=False,
)
del _behavior_probe
_registry = HandlerRegistry()


def get_registry() -> HandlerRegistry:
    """Get the global handler registry."""
    return _registry
