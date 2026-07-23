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


def filter_run_end_encoded(
    array: pa.RunEndEncodedArray,
    keep: pa.BooleanArray,
) -> pa.RunEndEncodedArray:
    """Filter a bounded REE slice while retaining its recursive encoded type."""
    if len(array) != len(keep):
        raise ValueError("REE filter mask length must match the logical slice")
    encoded_logical_type(array.type)
    safe_keep = cast("pa.BooleanArray", pc.fill_null(keep, False))
    return _filter_run_end_encoded(array, safe_keep, 0, [MAX_ENCODED_NODES])


def trim_run_end_encoded(array: pa.RunEndEncodedArray) -> pa.RunEndEncodedArray:
    """Return only intersecting runs, rebased to the current logical slice."""
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
) -> pa.Array:
    _consume_projection_budget(array, depth, budget)
    if isinstance(array, pa.DictionaryArray):
        dictionary_view = _project_encoded(
            array.dictionary,
            leaf_projection,
            depth + 1,
            budget,
        )
        return cast("pa.Array", pc.take(dictionary_view, array.indices))
    if isinstance(array, pa.RunEndEncodedArray):
        return _project_run_end_encoded(
            array,
            leaf_projection,
            depth,
            budget,
        )
    return leaf_projection(array)


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
    if array.nbytes > MAX_ENCODED_VIEW_BYTES:
        raise ValueError("encoded array exceeds logical-view byte limit")
    budget[0] -= 1
