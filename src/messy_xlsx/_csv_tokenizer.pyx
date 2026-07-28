# cython: language_level=3

from cpython.mem cimport PyMem_Free, PyMem_Malloc, PyMem_Realloc
from libc.string cimport memcpy

import sys


NATIVE_API_VERSION = 1
PANDAS_SEMANTIC_VERSION = "3.0.5"
NATIVE_SOURCE_SHA256 = NATIVE_SOURCE_SHA256_VALUE


cdef class _NativeDebugState:
    cdef readonly object state

    def __cinit__(self, object state):
        self.state = state


def _exercise_allocation_path_for_tests(
    object payload,
    Py_ssize_t growth,
    bint fail_reallocation=False,
):
    """Exercise the checked PyMem allocation path without retaining ownership."""

    cdef const unsigned char[::1] view
    cdef Py_ssize_t payload_size
    cdef Py_ssize_t requested_size
    cdef Py_ssize_t maximum_size = sys.maxsize
    cdef size_t initial_capacity
    cdef size_t requested_capacity
    cdef unsigned char *allocation = NULL
    cdef unsigned char *resized = NULL

    try:
        view = payload
    except (TypeError, ValueError, BufferError) as exc:
        raise TypeError("payload must expose a contiguous byte buffer") from exc

    payload_size = view.shape[0]
    if growth < 0:
        raise ValueError("growth must be non-negative")
    if growth > maximum_size - payload_size:
        raise OverflowError("native allocation size overflow")

    requested_size = payload_size + growth
    initial_capacity = <size_t>(payload_size if payload_size > 0 else 1)
    requested_capacity = <size_t>(requested_size if requested_size > 0 else 1)
    allocation = <unsigned char *>PyMem_Malloc(initial_capacity)
    if allocation == NULL:
        raise MemoryError()

    try:
        if payload_size > 0:
            memcpy(allocation, &view[0], <size_t>payload_size)
        if not fail_reallocation:
            resized = <unsigned char *>PyMem_Realloc(allocation, requested_capacity)
        if resized == NULL:
            raise MemoryError()
        allocation = resized
        return (<char *>allocation)[:payload_size]
    finally:
        PyMem_Free(allocation)


cdef class NativeCSVTokenizer:
    """Fail-closed ABI shell; CSV parsing is deliberately not implemented yet."""

    cdef object _config
    cdef object _source
    cdef object _state
    cdef unsigned char *_allocation
    cdef Py_ssize_t _allocation_capacity

    def __cinit__(self):
        self._config = None
        self._source = None
        self._state = "new"
        self._allocation = NULL
        self._allocation_capacity = 0

    def __init__(self, object config):
        self._config = config

    property debug_state:
        def __get__(self):
            return _NativeDebugState(self._state)

    def bind(self, object source):
        if self._state != "new":
            raise RuntimeError("native CSV tokenizer bind is one-shot")
        self._source = source
        self._state = "bound"

    def read_batch(self, Py_ssize_t requested_rows, object on_warning):
        cdef const unsigned char[::1] view
        cdef object payload
        cdef Py_ssize_t payload_size
        cdef Py_ssize_t requested_capacity
        cdef unsigned char *resized = NULL
        cdef Py_ssize_t read_size = 64

        if self._state != "bound":
            raise RuntimeError("native CSV tokenizer is not bound")
        if requested_rows <= 0:
            raise ValueError("requested_rows must be positive")
        if not callable(on_warning):
            raise TypeError("on_warning must be callable")

        self._state = "reading"
        try:
            payload = self._source.read(read_size)
            if len(payload) > read_size:
                raise ValueError("source.read() returned more bytes than requested")
            try:
                view = payload
            except (TypeError, ValueError, BufferError) as exc:
                raise TypeError(
                    "source.read() must return a contiguous byte buffer"
                ) from exc

            payload_size = view.shape[0]
            if payload_size == sys.maxsize:
                raise OverflowError("native allocation size overflow")
            requested_capacity = payload_size + 1
            self._allocation = <unsigned char *>PyMem_Malloc(
                <size_t>(payload_size if payload_size > 0 else 1)
            )
            if self._allocation == NULL:
                raise MemoryError()
            self._allocation_capacity = payload_size if payload_size > 0 else 1
            if payload_size > 0:
                memcpy(self._allocation, &view[0], <size_t>payload_size)

            resized = <unsigned char *>PyMem_Realloc(
                self._allocation,
                <size_t>requested_capacity,
            )
            if resized == NULL:
                raise MemoryError()
            self._allocation = resized
            self._allocation_capacity = requested_capacity

            view = None
            on_warning({"kind": "abi_shell", "bytes_seen": payload_size})
            raise NotImplementedError("native CSV ABI shell does not parse CSV yet")
        except BaseException:
            self._state = "terminal"
            raise

    def close(self):
        if self._state == "reading":
            raise RuntimeError("cannot close native CSV tokenizer while reading")
        if self._state == "closed":
            return
        if self._allocation != NULL:
            PyMem_Free(self._allocation)
            self._allocation = NULL
            self._allocation_capacity = 0
        self._source = None
        self._config = None
        self._state = "closed"

    def __dealloc__(self):
        if self._allocation != NULL:
            PyMem_Free(self._allocation)
            self._allocation = NULL
