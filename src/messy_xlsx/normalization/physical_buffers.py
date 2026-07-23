"""Bounded physical-buffer accounting shared by plan and runtime validation."""

from __future__ import annotations

from collections.abc import Iterator, Sequence
from typing import Final

import pyarrow as pa

MAX_PHYSICAL_BUFFER_DEPTH: Final = 32
MAX_PHYSICAL_BUFFER_NODES: Final = 1_024
MAX_PHYSICAL_BUFFER_ENTRIES: Final = 1_024
MAX_BUFFER_PARENT_NODES: Final = 1_024


class PhysicalBufferTraversalError(ValueError):
    """Report that physical-buffer accounting exceeded a structural bound."""


def unique_physical_buffer_bytes(
    arrays: Sequence[pa.Array],
    *,
    max_depth: int = MAX_PHYSICAL_BUFFER_DEPTH,
    max_nodes: int = MAX_PHYSICAL_BUFFER_NODES,
    max_buffer_entries: int = MAX_PHYSICAL_BUFFER_ENTRIES,
    max_parent_nodes: int = MAX_BUFFER_PARENT_NODES,
) -> int:
    """Return deduplicated root-buffer bytes after bounded structural preflight."""
    type_budget = [max_nodes]
    buffer_entry_budget = [max_buffer_entries]
    for array in arrays:
        _preflight_physical_type(
            array.type,
            0,
            type_budget,
            buffer_entry_budget,
            max_depth,
        )

    seen_buffers: set[tuple[int, int]] = set()
    array_budget = [max_nodes]
    return sum(
        _unique_array_buffer_bytes(
            array,
            seen_buffers,
            0,
            array_budget,
            max_depth,
            max_parent_nodes,
        )
        for array in arrays
    )


def preflight_physical_type(
    value_type: pa.DataType,
    *,
    max_depth: int = MAX_PHYSICAL_BUFFER_DEPTH,
    max_nodes: int = MAX_PHYSICAL_BUFFER_NODES,
    max_buffer_entries: int = MAX_PHYSICAL_BUFFER_ENTRIES,
) -> None:
    """Reject types whose flattened buffer enumeration cannot be bounded."""
    _preflight_physical_type(
        value_type,
        0,
        [max_nodes],
        [max_buffer_entries],
        max_depth,
    )


def _preflight_physical_type(
    value_type: pa.DataType,
    depth: int,
    node_budget: list[int],
    buffer_entry_budget: list[int],
    max_depth: int,
) -> None:
    _consume_structural_budget(depth, node_budget, max_depth)
    if value_type.has_variadic_buffers:
        raise PhysicalBufferTraversalError
    if buffer_entry_budget[0] < value_type.num_buffers:
        raise PhysicalBufferTraversalError
    buffer_entry_budget[0] -= value_type.num_buffers
    for child_type in _physical_child_types(value_type):
        _preflight_physical_type(
            child_type,
            depth + 1,
            node_budget,
            buffer_entry_budget,
            max_depth,
        )


def _unique_array_buffer_bytes(
    array: pa.Array,
    seen_buffers: set[tuple[int, int]],
    depth: int,
    budget: list[int],
    max_depth: int,
    max_parent_nodes: int,
) -> int:
    _consume_structural_budget(depth, budget, max_depth)
    total = _unique_flattened_buffer_bytes(
        array,
        seen_buffers,
        max_parent_nodes,
    )
    for child in _physical_child_arrays(array):
        total += _unique_array_buffer_bytes(
            child,
            seen_buffers,
            depth + 1,
            budget,
            max_depth,
            max_parent_nodes,
        )
    return total


def _consume_structural_budget(
    depth: int,
    budget: list[int],
    max_depth: int,
) -> None:
    if depth > max_depth or budget[0] < 1:
        raise PhysicalBufferTraversalError
    budget[0] -= 1


def _unique_flattened_buffer_bytes(
    array: pa.Array,
    seen_buffers: set[tuple[int, int]],
    max_parent_nodes: int,
) -> int:
    total = 0
    for buffer in array.buffers():
        if buffer is None:
            continue
        root = _bounded_root_buffer(buffer, max_parent_nodes)
        identity = (root.address, root.size)
        if identity not in seen_buffers:
            seen_buffers.add(identity)
            total += root.size
    return total


def _bounded_root_buffer(buffer: pa.Buffer, max_parent_nodes: int) -> pa.Buffer:
    retained: list[pa.Buffer] = []
    seen_wrappers: set[int] = set()
    current = buffer
    for _ in range(max_parent_nodes):
        wrapper_identity = id(current)
        if wrapper_identity in seen_wrappers:
            raise PhysicalBufferTraversalError
        retained.append(current)
        seen_wrappers.add(wrapper_identity)
        parent = current.parent
        if parent is None:
            return current
        current = parent
    raise PhysicalBufferTraversalError


def _physical_child_types(value_type: pa.DataType) -> Iterator[pa.DataType]:
    if pa.types.is_dictionary(value_type):
        yield value_type.value_type
    elif isinstance(value_type, pa.BaseExtensionType):
        yield value_type.storage_type
    elif (
        pa.types.is_list(value_type)
        or pa.types.is_large_list(value_type)
        or pa.types.is_fixed_size_list(value_type)
        or pa.types.is_list_view(value_type)
        or pa.types.is_large_list_view(value_type)
    ):
        yield value_type.value_type
    elif pa.types.is_map(value_type):
        yield value_type.field(0).type
    elif pa.types.is_struct(value_type) or pa.types.is_union(value_type):
        for index in range(value_type.num_fields):
            yield value_type.field(index).type
    elif pa.types.is_run_end_encoded(value_type):
        yield value_type.run_end_type
        yield value_type.value_type


def _physical_child_arrays(array: pa.Array) -> Iterator[pa.Array]:
    if isinstance(array, pa.DictionaryArray):
        yield array.dictionary
    elif isinstance(array, pa.ExtensionArray):
        yield array.storage
    elif (
        pa.types.is_list(array.type)
        or pa.types.is_large_list(array.type)
        or pa.types.is_fixed_size_list(array.type)
        or pa.types.is_list_view(array.type)
        or pa.types.is_large_list_view(array.type)
        or pa.types.is_map(array.type)
    ):
        yield array.values
    elif pa.types.is_struct(array.type) or pa.types.is_union(array.type):
        for index in range(array.type.num_fields):
            yield array.field(index)
    elif isinstance(array, pa.RunEndEncodedArray):
        yield array.run_ends
        yield array.values
