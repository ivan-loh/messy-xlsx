from __future__ import annotations

import gc
import importlib
import sys
from hashlib import sha256
from pathlib import Path
from typing import Any

import pytest

PYX_PATH = Path(__file__).resolve().parents[2] / "src" / "messy_xlsx" / "_csv_tokenizer.pyx"


def test_native_module_handshake_and_initial_state() -> None:
    native = importlib.import_module("messy_xlsx._csv_tokenizer")

    assert native.NATIVE_API_VERSION == 1
    assert native.PANDAS_SEMANTIC_VERSION == "3.0.5"
    assert sha256(PYX_PATH.read_bytes()).hexdigest() == native.NATIVE_SOURCE_SHA256
    tokenizer = native.NativeCSVTokenizer(object())
    assert tokenizer.debug_state.state == "new"
    tokenizer.close()
    tokenizer.close()
    assert tokenizer.debug_state.state == "closed"


def test_constructor_and_bind_are_inert_until_read() -> None:
    native = importlib.import_module("messy_xlsx._csv_tokenizer")

    class Source:
        def read(self, size: int) -> bytes:
            raise AssertionError(f"inert shell unexpectedly read {size} bytes")

    tokenizer = native.NativeCSVTokenizer(object())
    tokenizer.bind(Source())
    assert tokenizer.debug_state.state == "bound"
    tokenizer.close()


def test_shell_read_exercises_memoryview_source_callback_and_terminal_state() -> None:
    native = importlib.import_module("messy_xlsx._csv_tokenizer")
    read_sizes: list[int] = []
    callbacks: list[Any] = []

    class Source:
        def read(self, size: int) -> memoryview:
            assert 1 <= size <= 64 * 1024
            read_sizes.append(size)
            return memoryview(b"abi-shell")

    tokenizer = native.NativeCSVTokenizer(object())
    tokenizer.bind(Source())

    with pytest.raises(NotImplementedError, match="ABI shell"):
        tokenizer.read_batch(1, callbacks.append)

    assert len(read_sizes) == 1
    assert len(callbacks) == 1
    assert tokenizer.debug_state.state == "terminal"
    tokenizer.close()
    assert tokenizer.debug_state.state == "closed"


def test_close_is_rejected_reentrantly_from_source_read() -> None:
    native = importlib.import_module("messy_xlsx._csv_tokenizer")
    observed_states: list[str] = []
    callbacks: list[Any] = []
    tokenizer = native.NativeCSVTokenizer(object())

    class Source:
        def read(self, size: int) -> bytes:
            assert size > 0
            with pytest.raises(RuntimeError, match="while reading"):
                tokenizer.close()
            observed_states.append(tokenizer.debug_state.state)
            return b"still-owned"

    tokenizer.bind(Source())
    with pytest.raises(NotImplementedError, match="ABI shell"):
        tokenizer.read_batch(1, callbacks.append)

    assert observed_states == ["reading"]
    assert callbacks == [{"kind": "abi_shell", "bytes_seen": 11}]
    assert tokenizer.debug_state.state == "terminal"
    tokenizer.close()
    assert tokenizer.debug_state.state == "closed"


def test_close_is_rejected_reentrantly_from_warning_callback() -> None:
    native = importlib.import_module("messy_xlsx._csv_tokenizer")
    callback_payloads: list[Any] = []
    observed_states: list[str] = []
    tokenizer = native.NativeCSVTokenizer(object())

    class Source:
        def read(self, size: int) -> bytes:
            assert size > 0
            return b"still-owned"

    def on_warning(payload: Any) -> None:
        callback_payloads.append(payload)
        with pytest.raises(RuntimeError, match="while reading"):
            tokenizer.close()
        observed_states.append(tokenizer.debug_state.state)

    tokenizer.bind(Source())
    with pytest.raises(NotImplementedError, match="ABI shell"):
        tokenizer.read_batch(1, on_warning)

    assert callback_payloads == [{"kind": "abi_shell", "bytes_seen": 11}]
    assert observed_states == ["reading"]
    assert tokenizer.debug_state.state == "terminal"
    tokenizer.close()
    assert tokenizer.debug_state.state == "closed"


def test_allocation_reallocation_success_and_fault_paths() -> None:
    native = importlib.import_module("messy_xlsx._csv_tokenizer")

    assert native._exercise_allocation_path_for_tests(memoryview(b"abi"), 29, False) == b"abi"
    with pytest.raises(MemoryError):
        native._exercise_allocation_path_for_tests(memoryview(b"abi"), 29, True)
    with pytest.raises(OverflowError):
        native._exercise_allocation_path_for_tests(b"x", sys.maxsize, False)


def test_lifecycle_construction_and_close_stress() -> None:
    native = importlib.import_module("messy_xlsx._csv_tokenizer")

    for index in range(1_000):
        tokenizer = native.NativeCSVTokenizer(index)
        if index % 2 == 0:
            tokenizer.close()
            tokenizer.close()
    del tokenizer
    gc.collect()
