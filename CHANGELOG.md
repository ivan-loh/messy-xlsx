# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.8.0] - 2026-02-25

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
