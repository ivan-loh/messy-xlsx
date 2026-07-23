"""Bounded logical projections for nested Arrow dictionary and REE arrays."""

from __future__ import annotations

from collections.abc import Callable
from typing import Final, cast

import pyarrow as pa
import pyarrow.compute as pc

MAX_ENCODED_DEPTH: Final = 32
MAX_ENCODED_NODES: Final = 1_024
MAX_ENCODED_VIEW_BYTES: Final = 8 * 1024 * 1024


def encoded_logical_type(value_type: pa.DataType) -> pa.DataType:
    """Return the leaf type beneath a bounded dictionary/REE wrapper chain."""
    depth = 0
    nodes = MAX_ENCODED_NODES
    while pa.types.is_dictionary(value_type) or pa.types.is_run_end_encoded(value_type):
        if depth >= MAX_ENCODED_DEPTH or nodes < 1:
            raise ValueError("encoded array exceeds logical-view structural limits")
        value_type = value_type.value_type
        depth += 1
        nodes -= 1
    return value_type


def encoded_logical_view(array: pa.Array) -> pa.Array:
    """Decode nested dictionary/REE wrappers into the bounded logical slice."""
    return _project_encoded(array, lambda leaf: leaf, 0, [MAX_ENCODED_NODES])


def encoded_logical_validity(array: pa.Array) -> pa.BooleanArray:
    """Project logical validity through arbitrary dictionary/REE wrapper order."""
    projected = _project_encoded(
        array,
        lambda leaf: cast("pa.Array", pc.is_valid(leaf)),
        0,
        [MAX_ENCODED_NODES],
    )
    return cast("pa.BooleanArray", pc.fill_null(projected, False))


def encoded_has_no_logical_nulls(array: pa.Array) -> bool:
    """Prove logical density from encoded metadata and bounded references."""
    return _has_no_logical_nulls(array, 0, [MAX_ENCODED_NODES])


def filter_encoded(array: pa.Array, keep: pa.BooleanArray) -> pa.Array:
    """Filter an array while retaining any recursive dictionary/REE type."""
    if len(array) != len(keep):
        raise ValueError("encoded filter mask length must match the logical slice")
    if isinstance(array, (pa.DictionaryArray, pa.RunEndEncodedArray)):
        _validate_physical_buffer_budget(array)
    safe_keep = cast("pa.BooleanArray", pc.fill_null(keep, False))
    if pc.all(safe_keep).as_py() is True:
        return array
    if isinstance(array, pa.RunEndEncodedArray):
        return filter_run_end_encoded(array, safe_keep)
    return cast("pa.Array", pc.filter(array, safe_keep))


def mask_encoded(array: pa.Array, mask: pa.BooleanArray) -> pa.Array:
    """Replace selected logical values with null while retaining encoded type."""
    if len(array) != len(mask):
        raise ValueError("encoded mask length must match the logical slice")
    if isinstance(array, (pa.DictionaryArray, pa.RunEndEncodedArray)):
        _validate_physical_buffer_budget(array)
    safe_mask = cast("pa.BooleanArray", pc.fill_null(mask, False))
    if pc.any(safe_mask).as_py() is not True:
        return array
    if isinstance(array, pa.RunEndEncodedArray):
        encoded_logical_type(array.type)
        return _mask_run_end_encoded(array, safe_mask, 0, [MAX_ENCODED_NODES])
    if isinstance(array, pa.DictionaryArray):
        indices = cast(
            "pa.Array",
            pc.if_else(
                safe_mask,
                pa.nulls(len(array), type=array.indices.type),
                array.indices,
            ),
        )
        return pa.DictionaryArray.from_arrays(
            indices,
            array.dictionary,
            ordered=cast("pa.DictionaryType", array.type).ordered,
        )
    return cast(
        "pa.Array",
        pc.if_else(
            safe_mask,
            pa.nulls(len(array), type=array.type),
            array,
        ),
    )


def filter_run_end_encoded(
    array: pa.RunEndEncodedArray,
    keep: pa.BooleanArray,
) -> pa.RunEndEncodedArray:
    """Filter a bounded REE slice while retaining its recursive encoded type."""
    if len(array) != len(keep):
        raise ValueError("REE filter mask length must match the logical slice")
    _validate_physical_buffer_budget(array)
    encoded_logical_type(array.type)
    safe_keep = cast("pa.BooleanArray", pc.fill_null(keep, False))
    return _filter_run_end_encoded(array, safe_keep, 0, [MAX_ENCODED_NODES])


def trim_run_end_encoded(array: pa.RunEndEncodedArray) -> pa.RunEndEncodedArray:
    """Return only intersecting runs, rebased to the current logical slice."""
    _validate_physical_buffer_budget(array)
    physical_offset = array.find_physical_offset()
    physical_length = array.find_physical_length()
    run_values = array.values.slice(physical_offset, physical_length)
    intersecting_run_ends = array.run_ends.slice(physical_offset, physical_length)
    logical_stop = array.offset + len(array)
    rebased_run_ends = pa.array(
        [
            min(cast("int", run_end.as_py()), logical_stop) - array.offset
            for run_end in intersecting_run_ends
        ],
        type=intersecting_run_ends.type,
    )
    return pa.RunEndEncodedArray.from_arrays(rebased_run_ends, run_values)


def _project_encoded(
    array: pa.Array,
    leaf_projection: Callable[[pa.Array], pa.Array],
    depth: int,
    budget: list[int],
    indices: pa.Array | None = None,
) -> pa.Array:
    _consume_projection_budget(array, depth, budget)
    if indices is not None:
        unique_indices, restore_indices = _compact_take_indices(indices, len(array))
        projected_unique = _project_unique_encoded(
            array,
            unique_indices,
            leaf_projection,
            depth,
            budget,
        )
        return cast("pa.Array", pc.take(projected_unique, restore_indices))
    if isinstance(array, pa.DictionaryArray):
        return _project_encoded(
            array.dictionary,
            leaf_projection,
            depth + 1,
            budget,
            array.indices,
        )
    if isinstance(array, pa.RunEndEncodedArray):
        return _project_run_end_encoded(
            array,
            leaf_projection,
            depth,
            budget,
        )
    return leaf_projection(array)


def _has_no_logical_nulls(
    array: pa.Array,
    depth: int,
    budget: list[int],
    indices: pa.Array | None = None,
) -> bool:
    _consume_projection_budget(array, depth, budget)
    if indices is not None:
        if indices.null_count:
            return False
        unique_indices = cast("pa.Array", pc.unique(indices))
        if not len(unique_indices):
            return True
        return _selected_values_have_no_logical_nulls(
            array,
            unique_indices,
            depth,
            budget,
        )
    if isinstance(array, pa.DictionaryArray):
        if array.null_count:
            return False
        if _metadata_proves_all_values_valid(
            array.dictionary,
            depth + 1,
            budget,
        ):
            return True
        referenced = cast("pa.Array", pc.unique(array.indices))
        return _has_no_logical_nulls(
            array.dictionary,
            depth + 1,
            budget,
            referenced,
        )
    if isinstance(array, pa.RunEndEncodedArray):
        trimmed = trim_run_end_encoded(array)
        return _has_no_logical_nulls(
            trimmed.values,
            depth + 1,
            budget,
        )
    return bool(array.null_count == 0)


def _selected_values_have_no_logical_nulls(
    array: pa.Array,
    indices: pa.Array,
    depth: int,
    budget: list[int],
) -> bool:
    if isinstance(array, pa.DictionaryArray):
        dictionary_indices = cast("pa.Array", pc.take(array.indices, indices))
        if dictionary_indices.null_count:
            return False
        if _metadata_proves_all_values_valid(
            array.dictionary,
            depth + 1,
            budget,
        ):
            return True
        referenced = cast("pa.Array", pc.unique(dictionary_indices))
        return _has_no_logical_nulls(
            array.dictionary,
            depth + 1,
            budget,
            referenced,
        )
    if isinstance(array, pa.RunEndEncodedArray):
        physical_indices = cast(
            "pa.Array",
            pc.unique(_run_end_physical_indices(array, indices)),
        )
        return _has_no_logical_nulls(
            array.values,
            depth + 1,
            budget,
            physical_indices,
        )
    if array.null_count == 0:
        return True
    selected = cast("pa.Array", pc.take(array, indices))
    return bool(selected.null_count == 0)


def _metadata_proves_all_values_valid(
    array: pa.Array,
    depth: int,
    budget: list[int],
) -> bool:
    _consume_projection_budget(array, depth, budget)
    if isinstance(array, pa.DictionaryArray):
        return bool(
            array.null_count == 0
            and _metadata_proves_all_values_valid(
                array.dictionary,
                depth + 1,
                budget,
            )
        )
    if isinstance(array, pa.RunEndEncodedArray):
        trimmed = trim_run_end_encoded(array)
        return _metadata_proves_all_values_valid(
            trimmed.values,
            depth + 1,
            budget,
        )
    return bool(array.null_count == 0)


def _project_unique_encoded(
    array: pa.Array,
    indices: pa.Array,
    leaf_projection: Callable[[pa.Array], pa.Array],
    depth: int,
    budget: list[int],
) -> pa.Array:
    if isinstance(array, pa.DictionaryArray):
        dictionary_indices = cast("pa.Array", pc.take(array.indices, indices))
        return _project_encoded(
            array.dictionary,
            leaf_projection,
            depth + 1,
            budget,
            dictionary_indices,
        )
    if isinstance(array, pa.RunEndEncodedArray):
        physical_indices = _run_end_physical_indices(array, indices)
        return _project_encoded(
            array.values,
            leaf_projection,
            depth + 1,
            budget,
            physical_indices,
        )
    selected = cast("pa.Array", pc.take(array, indices))
    return leaf_projection(selected)


def _compact_take_indices(
    indices: pa.Array,
    upper_bound: int,
) -> tuple[pa.Int64Array, pa.Int64Array]:
    unique: list[int] = []
    positions: dict[int, int] = {}
    restore: list[int | None] = []
    for scalar in indices:
        if not scalar.is_valid:
            restore.append(None)
            continue
        value = cast("int", scalar.as_py())
        if value < 0 or value >= upper_bound:
            raise IndexError("encoded logical index is out of bounds")
        position = positions.get(value)
        if position is None:
            position = len(unique)
            positions[value] = position
            unique.append(value)
        restore.append(position)
    return (
        pa.array(unique, type=pa.int64()),
        pa.array(restore, type=pa.int64()),
    )


def _run_end_physical_indices(
    array: pa.RunEndEncodedArray,
    indices: pa.Array,
) -> pa.Int64Array:
    return pa.array(
        [array.slice(cast("int", scalar.as_py()), 1).find_physical_offset() for scalar in indices],
        type=pa.int64(),
    )


def _project_run_end_encoded(
    array: pa.RunEndEncodedArray,
    leaf_projection: Callable[[pa.Array], pa.Array],
    depth: int,
    budget: list[int],
) -> pa.Array:
    trimmed = trim_run_end_encoded(array)
    projected_values = _project_encoded(
        trimmed.values,
        leaf_projection,
        depth + 1,
        budget,
    )
    if not len(array):
        return projected_values.slice(0, 0)
    encoded = pa.RunEndEncodedArray.from_arrays(trimmed.run_ends, projected_values)
    return cast("pa.Array", pc.run_end_decode(encoded))


def _filter_run_end_encoded(
    array: pa.RunEndEncodedArray,
    keep: pa.BooleanArray,
    depth: int,
    budget: list[int],
) -> pa.RunEndEncodedArray:
    _consume_projection_budget(array, depth, budget)
    trimmed = trim_run_end_encoded(array)
    physical_positions = _decode_physical_positions(trimmed)
    filtered_positions = cast("pa.Array", pc.filter(physical_positions, keep))
    return _rebuild_run_end_encoded(
        trimmed,
        filtered_positions,
        depth,
        budget,
    )


def _take_run_end_encoded(
    array: pa.RunEndEncodedArray,
    indices: pa.Array,
    depth: int,
    budget: list[int],
) -> pa.RunEndEncodedArray:
    _consume_projection_budget(array, depth, budget)
    trimmed = trim_run_end_encoded(array)
    physical_positions = _decode_physical_positions(trimmed)
    selected_positions = cast("pa.Array", pc.take(physical_positions, indices))
    return _rebuild_run_end_encoded(
        trimmed,
        selected_positions,
        depth,
        budget,
    )


def _mask_run_end_encoded(
    array: pa.RunEndEncodedArray,
    mask: pa.BooleanArray,
    depth: int,
    budget: list[int],
) -> pa.RunEndEncodedArray:
    _consume_projection_budget(array, depth, budget)
    trimmed = trim_run_end_encoded(array)
    physical_positions = _decode_physical_positions(trimmed)
    masked_positions = cast(
        "pa.Array",
        pc.if_else(
            mask,
            pa.nulls(len(mask), type=physical_positions.type),
            physical_positions,
        ),
    )
    return _rebuild_run_end_encoded(
        trimmed,
        masked_positions,
        depth,
        budget,
    )


def _decode_physical_positions(array: pa.RunEndEncodedArray) -> pa.Array:
    position_runs = pa.RunEndEncodedArray.from_arrays(
        array.run_ends,
        pa.array(range(len(array.values)), type=pa.int64()),
    )
    return cast("pa.Array", pc.run_end_decode(position_runs))


def _rebuild_run_end_encoded(
    trimmed: pa.RunEndEncodedArray,
    physical_positions: pa.Array,
    depth: int,
    budget: list[int],
) -> pa.RunEndEncodedArray:
    encoded_positions = pc.run_end_encode(
        physical_positions,
        run_end_type=trimmed.run_ends.type,
    )
    selected_values = _take_encoded(
        trimmed.values,
        encoded_positions.values,
        depth + 1,
        budget,
    )
    return pa.RunEndEncodedArray.from_arrays(
        encoded_positions.run_ends,
        selected_values,
    )


def _take_encoded(
    array: pa.Array,
    indices: pa.Array,
    depth: int,
    budget: list[int],
) -> pa.Array:
    if isinstance(array, pa.RunEndEncodedArray):
        return _take_run_end_encoded(array, indices, depth, budget)
    _consume_projection_budget(array, depth, budget)
    return cast("pa.Array", pc.take(array, indices))


def _consume_projection_budget(
    array: pa.Array,
    depth: int,
    budget: list[int],
) -> None:
    if depth > MAX_ENCODED_DEPTH or budget[0] < 1:
        raise ValueError("encoded array exceeds logical-view structural limits")
    _validate_physical_buffer_budget(array)
    budget[0] -= 1


def _validate_physical_buffer_budget(array: pa.Array) -> None:
    seen_buffers: set[tuple[int, int]] = set()
    buffer_bytes = _unique_physical_buffer_bytes(
        array,
        seen_buffers,
        0,
        [MAX_ENCODED_NODES],
    )
    if buffer_bytes > MAX_ENCODED_VIEW_BYTES:
        raise ValueError("encoded array exceeds logical-view byte limit")


def _unique_physical_buffer_bytes(
    array: pa.Array,
    seen_buffers: set[tuple[int, int]],
    depth: int,
    budget: list[int],
) -> int:
    if depth > MAX_ENCODED_DEPTH or budget[0] < 1:
        raise ValueError("encoded array exceeds logical-view structural limits")
    budget[0] -= 1
    total = _unique_root_buffer_bytes(array, seen_buffers)
    for child in _physical_child_arrays(array):
        total += _unique_physical_buffer_bytes(
            child,
            seen_buffers,
            depth + 1,
            budget,
        )
    return total


def _unique_root_buffer_bytes(
    array: pa.Array,
    seen_buffers: set[tuple[int, int]],
) -> int:
    total = 0
    for original_buffer in array.buffers():
        if original_buffer is None:
            continue
        buffer = original_buffer
        while buffer.parent is not None:
            buffer = buffer.parent
        identity = (buffer.address, buffer.size)
        if identity not in seen_buffers:
            seen_buffers.add(identity)
            total += buffer.size
    return total


def _physical_child_arrays(array: pa.Array) -> tuple[pa.Array, ...]:
    if isinstance(array, pa.DictionaryArray):
        return (array.dictionary,)
    if isinstance(array, pa.ExtensionArray):
        return (array.storage,)
    if (
        pa.types.is_list(array.type)
        or pa.types.is_large_list(array.type)
        or pa.types.is_fixed_size_list(array.type)
        or pa.types.is_list_view(array.type)
        or pa.types.is_large_list_view(array.type)
        or pa.types.is_map(array.type)
    ):
        return (array.values,)
    if pa.types.is_struct(array.type) or pa.types.is_union(array.type):
        return tuple(array.field(index) for index in range(array.type.num_fields))
    if isinstance(array, pa.RunEndEncodedArray):
        return (array.run_ends, array.values)
    return ()
