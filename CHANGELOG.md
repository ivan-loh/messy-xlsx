# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.10.0] - 2026-07-21

### Changed

- Added `xlwt` to development dependencies so legacy `.xls` tests run instead of skip.
- Declared `pyarrow` as a runtime dependency because the `fastexcel` DataFrame conversion
  path requires it in clean installations.
- Reused the core `MessyWorkbook` pipeline for multi-sheet parsing, removing duplicate
  header detection, normalization, type coercion, and column cleaning code.
- Enabled the `fastexcel` path when structure analysis confirms that merged/hidden-cell
  handling is unnecessary.
- `SheetConfig.evaluate_formulas=False` now preserves formula expressions in DataFrames;
  the default reads cached results.
- Compiled sheet configuration, structure evidence, backend hints, and normalization
  choices into one private frozen parse plan before parsing.
- Centralized path, seekable-buffer, and read-once-stream access behind one internal source
  boundary shared by detection, analysis, validation, and format handlers.
- Reduced the source distribution to the package source and public project metadata instead
  of bundling test workbooks and internal development files.
- Gated PyPI publishing on strict tag/version validation, the complete reusable CI workflow,
  and the exact wheel and source distribution artifacts verified by that workflow.
- Made documentation builds and dependency vulnerability scans blocking CI checks and pinned
  release workflow actions to immutable commits.
- Preserved legacy custom detectors, handlers, and registry subclasses through fresh raw
  source borrows unless an extension explicitly opts into the internal handle contract.

### Fixed

- Exported `read_excel`, `read_excel_tables`, and `analyze_structure` through `__all__`.
- Added `SheetInfo.column_count` as a descriptive alias for `col_count`.
- Prevented `MessyTable.to_dataframe()` from mutating caller-owned configuration.
- Closed convenience-function workbooks and transient XLS/XLSX/CSV resources
  deterministically on both success and failure without closing caller-owned buffers.
- Made sheet structure access honor workbook header patterns and made table parsing inherit
  workbook configuration when no table-specific configuration is supplied.
- Prevented hidden rows outside the data region and sheet-global footer evidence from
  truncating detected table ranges.
- Restored caller-owned seekable streams to their exact entry position after success or
  failure, snapshotted non-seekable streams once, and retained caller ownership throughout.
- Normalized `bytearray` and `memoryview` streams only for backends that require true bytes,
  avoiding unconditional copies for ordinary seekable sources.
- Kept cached-value structure analysis separate from formula-expression detection so an
  uncached formula cannot change data bounds, header scoring, tables, or locale evidence.
- Limited CSV stream validation to the same bounded 1 KiB probe used for path inputs.
- Synchronized documented runtime dependencies with `pyproject.toml`.
- Corrected the property-test missing-value contract for `nil`/`NIL`.

## [0.9.0] - 2026-04-05

### Changed

- Raised minimum dependency versions to fastexcel 0.19.0, openpyxl 3.1.5, pandas 3.0.0,
  and NumPy 2.4.0.
- Updated optional formula/XLS and development dependency floors for the supported
  Python 3.11–3.14 matrix.
- Added 31 messy-workbook regression fixtures covering generated, malformed, and
  real-world-like inputs.

### Fixed

- Prevented numeric account codes and amounts from being misclassified as dates.
- Normalized mixed-locale numeric columns per value, preserving decimal values such as
  `0,22` as `0.22`.
- Preserved already-numeric values instead of re-parsing them through locale separators.
- Detected decimal separators without requiring thousands-separator evidence.
- Accounted for hidden rows when applying `skip_rows` during header detection.
- Restricted footer detection to specific footer phrases instead of matching bare `total`
  or `sum` text.
- Scanned across blank rows to detect separated footers.
- Constrained parsing of multi-table sheets to the detected primary table.
- Prevented the CSV fallback from decoding corrupted XLSX files as text.
- Improved multi-row header selection with an underscore-based tie-breaker.

## [0.8.0] - 2026-04-04

### Added

- **Enum types** for all configuration constants (`MergeStrategy`, `HeaderDetectionMode`,
  `HeaderFallback`, `DataType`, `FormatType`) with backward-compatible `StrEnum` base
  — raw strings like `"fill"` still work everywhere.
- **Input validation** on `SheetConfig` and `FormulaConfig` dataclasses via `__post_init__`:
  - `skip_rows`, `header_rows`, `skip_footer` must be >= 0
  - `header_confidence_threshold` must be 0.0–1.0
  - `FormulaConfig.max_iterations` and `max_depth` must be >= 1
- **PEP 561 `py.typed` marker** for downstream type-checker support.
- **CI workflow hardening**:
  - Separate lint, test, security, and build jobs.
  - Cross-platform matrix (Ubuntu, macOS, Windows) × Python 3.10–3.13.
  - Coverage gate: `--cov-fail-under=75`.
  - Security scanning via `bandit` + `pip-audit`.
- **Dependabot** configuration for automated dependency updates (pip + GitHub Actions).
- **Pre-commit hooks** (ruff check, ruff format, trailing-whitespace, check-yaml/toml).
- **Makefile** with `install`, `test`, `lint`, `format`, `typecheck`, `ci`, `benchmark`,
  `docs`, and `clean` targets.
- **Benchmark tests** using `pytest-benchmark` for regression tracking.
- **Documentation site** scaffolding with MkDocs Material + mkdocstrings.

### Changed

- **Expanded ruff rules**: added `SIM` (simplify), `RET` (return), `RUF` (ruff-specific),
  `C901` (McCabe complexity ≤ 10) to lint configuration.
- **Shared handler helpers**: extracted `is_fileobj()`, `reset_buffer()`, `get_file_desc()`,
  `read_file_content()` into `parsing/base_handler.py` — removed duplicates from
  `xlsx_handler.py`, `csv_handler.py`, `xls_handler.py`, and `handler_registry.py`.
- **Narrowed exception handling** across the codebase:
  - `workbook.py`: formula load/eval failures now catch specific types
    (`OSError`, `ValueError`, `TypeError`, `KeyError`) and emit `logger.debug()`.
  - `csv_handler.py`: metadata detection catches `ParserError`, `EmptyDataError`,
    `UnicodeDecodeError`, `ValueError`, `OSError` with debug logging.
  - `locale_detector.py`: cell access catches `AttributeError`, `IndexError`, `TypeError`.
  - `xlsx_handler.py`: fastexcel fallback catches `OSError`, `ValueError`, `RuntimeError`.
  - `_is_cell_merged` / `_is_cell_hidden`: narrowed to `AttributeError`, `TypeError`.
- **Replaced bare `assert`** in `get_cell()` with proper `FileError` raise.
- **Fixed `__version__`** from `"0.1.0"` to `"0.7.2"` (now `"0.8.0"`).

### Fixed

- Silent `except Exception: pass` blocks that swallowed errors without logging.
- Bare `assert self._wb is not None` that would raise `AssertionError` instead of
  a meaningful `FileError` when workbook was not loaded.
