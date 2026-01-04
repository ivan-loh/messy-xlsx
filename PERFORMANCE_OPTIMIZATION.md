# Performance Optimization Plan for messy-xlsx

## Problem Analysis

**Current bottleneck:** Structure analysis for large files (>1MB)

### Profiling Results
- **customers.xlsx** (2.5MB): 15+ minutes ❌
- **accounts_receivable.xlsx** (23KB): < 1 second ✅
- **cost_analysis.xlsx** (81KB, 660 rows): < 2 seconds ✅

### Root Causes

| Function | Complexity | Bottleneck |
|----------|-----------|------------|
| `_detect_data_region()` | O(rows × cols) | **Iterates EVERY cell** |
| `_detect_blank_rows()` | O(rows × cols) | Checks every cell for blanks |
| `_detect_multiple_tables()` | O(rows × cols) | Calls _detect_blank_rows() |
| `_detect_headers()` | O(cols) × 10 rows | Acceptable (limited to 10 rows) |
| `_detect_formulas()` | O(sample_size) | Already optimized (samples 100 cells) |

**Critical Issue:** Line 141-152 in `structure_analyzer.py`
```python
for row_idx, row in enumerate(ws.iter_rows(), start=1):  # ← Reads ALL rows!
    for col_idx, cell in enumerate(row, start=1):        # ← Reads ALL cells!
```

For 50MB file with 100,000 rows × 50 columns = **5,000,000 cell reads!**

---

## Optimization Strategy

### Strategy 1: **Sampling-Based Analysis** (Fastest)

**Concept:** Analyze only a sample of rows to detect structure

```python
class StructureAnalyzer:
    # Configuration
    MAX_ROWS_FULL_ANALYSIS = 10_000     # Full analysis threshold
    SAMPLE_SIZE_LARGE = 5_000            # Sample size for large files
    SAMPLE_INTERVAL = 10                 # Sample every Nth row

    def analyze(self, file_path, sheet, force=False):
        # Get sheet dimensions fast
        ws = wb[sheet]
        total_rows = ws.max_row
        total_cols = ws.max_column

        if total_rows <= self.MAX_ROWS_FULL_ANALYSIS:
            # Small file: full analysis
            return self._analyze_full(ws)
        else:
            # Large file: sampled analysis
            return self._analyze_sampled(ws, total_rows, total_cols)
```

**Implementation:**

```python
def _detect_data_region_fast(self, ws, max_row, max_col) -> dict:
    """Fast data region detection using ws.max_row/max_column."""
    # Use openpyxl's built-in dimension calculation
    # Then verify with spot checks

    # Start: Find first non-empty row
    start_row = None
    for row in range(1, min(100, max_row + 1)):
        for col in range(1, max_col + 1):
            if ws.cell(row, col).value is not None:
                start_row = row
                break
        if start_row:
            break

    # End: Use max_row but verify last few rows
    end_row = max_row
    for row in range(max_row, max(1, max_row - 50), -1):
        has_data = any(
            ws.cell(row, col).value is not None
            for col in range(1, min(20, max_col + 1))
        )
        if has_data:
            end_row = row
            break

    # Columns: Check first 100 rows only
    min_col, max_col_actual = None, None
    for row in range(start_row or 1, min((start_row or 1) + 100, end_row + 1)):
        for col in range(1, max_col + 1):
            if ws.cell(row, col).value is not None:
                if min_col is None or col < min_col:
                    min_col = col
                if max_col_actual is None or col > max_col_actual:
                    max_col_actual = col

    return {
        "start_row": start_row or 1,
        "end_row": end_row,
        "start_col": min_col or 1,
        "end_col": max_col_actual or max_col,
    }
```

**Time Complexity:** O(1) for dimensions + O(100 × cols) for verification
**Speedup:** 100-1000x for large files

---

### Strategy 2: **Progressive Loading** (Memory Efficient)

**Concept:** Load data in chunks, analyze incrementally

```python
def _analyze_with_chunks(self, ws, total_rows, chunk_size=1000):
    """Analyze file in chunks to limit memory usage."""

    # Phase 1: Analyze first chunk (headers, metadata)
    first_chunk = self._read_chunk(ws, 1, chunk_size)
    header_info = self._detect_headers_from_chunk(first_chunk)

    # Phase 2: Sample middle chunks for patterns
    if total_rows > chunk_size * 2:
        middle_chunk = self._read_chunk(ws, total_rows // 2, chunk_size)
        # Verify header pattern consistency

    # Phase 3: Analyze last chunk (footers)
    if total_rows > chunk_size:
        last_chunk = self._read_chunk(ws, total_rows - chunk_size, chunk_size)
        footer_info = self._detect_footers_from_chunk(last_chunk)

    return self._merge_chunk_results(...)
```

**Benefits:**
- Constant memory usage
- Can handle 100MB+ files
- Still gets accurate structure detection

---

### Strategy 3: **Fast Path for Simple Files** (Skip Analysis)

**Concept:** Detect "clean" files and skip expensive analysis

```python
def _is_simple_file(self, ws, max_row, max_col) -> bool:
    """Detect if file has simple structure (skip analysis)."""

    # Heuristics for simple files:
    checks = [
        max_row < 100_000,                    # Not too large
        max_col < 100,                        # Not too wide
        len(ws.merged_cells.ranges) == 0,     # No merged cells
        len(ws.row_dimensions) < 10,          # No custom row formatting
        len(ws.column_dimensions) < 10,       # No custom col formatting
    ]

    if all(checks):
        # Assume: row 1 = headers, row 2+ = data
        return True

    return False

def analyze(self, file_path, sheet, force=False):
    ws = wb[sheet]

    if self._is_simple_file(ws, ws.max_row, ws.max_column):
        # Fast path: return default structure
        return StructureInfo(
            data_start_row=1,
            data_end_row=ws.max_row,
            data_start_col=1,
            data_end_col=ws.max_column,
            header_row=1,
            header_confidence=0.95,
            # ... defaults
        )

    # Complex file: full analysis
    return self._analyze_full(ws)
```

**Speedup:** 1000x for simple files (instant vs minutes)

---

### Strategy 4: **Parallel Processing** (Multi-threading)

**Concept:** Process multiple sheets in parallel

```python
from concurrent.futures import ThreadPoolExecutor, as_completed

class MessyWorkbook:
    def to_dataframes(self, config=None, max_workers=4):
        """Parse all sheets in parallel."""

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(self._parse_sheet, name, config): name
                for name in self._sheet_names
            }

            results = {}
            for future in as_completed(futures):
                sheet_name = futures[future]
                try:
                    results[sheet_name] = future.result()
                except Exception as e:
                    # Log error, continue
                    pass

            return results
```

**Use Cases:**
- Multi-sheet workbooks (10+ sheets)
- Batch processing of multiple files

**Speedup:** Near-linear with CPU cores (4x on 4-core)

---

### Strategy 5: **Lazy Loading** (On-Demand Analysis)

**Concept:** Don't analyze until explicitly needed

```python
class MessyWorkbook:
    def __init__(self, file_path, lazy=True):
        self._lazy = lazy
        self._structures = {}  # Cache per sheet

        if not lazy:
            # Pre-analyze all sheets
            for sheet in self.sheet_names:
                self._structures[sheet] = self._analyze_structure(sheet)

    def get_structure(self, sheet=None):
        """Analyze on first access only."""
        sheet = sheet or self.sheet_names[0]

        if sheet not in self._structures:
            self._structures[sheet] = self._analyze_structure(sheet)

        return self._structures[sheet]
```

**Benefits:**
- Fast initialization
- Only pay for what you use
- Good for single-sheet access

---

### Strategy 6: **Smart Sampling** (Statistical Approach)

**Concept:** Sample strategically to get accurate results with minimal reads

```python
def _detect_blank_rows_sampled(self, ws, data_region, sample_size=1000):
    """Detect blank rows by sampling."""
    total_rows = data_region["end_row"] - data_region["start_row"] + 1

    if total_rows <= sample_size:
        # Small enough: check all rows
        return self._detect_blank_rows(ws, data_region)

    # Sample: first 200, last 200, random 600 in middle
    sample_rows = []
    sample_rows.extend(range(data_region["start_row"], data_region["start_row"] + 200))
    sample_rows.extend(range(data_region["end_row"] - 200, data_region["end_row"] + 1))

    # Random middle samples
    import random
    middle_start = data_region["start_row"] + 200
    middle_end = data_region["end_row"] - 200
    if middle_end > middle_start:
        middle_samples = random.sample(
            range(middle_start, middle_end),
            min(600, middle_end - middle_start)
        )
        sample_rows.extend(middle_samples)

    # Check sampled rows
    blank_rows = []
    for row_idx in sorted(set(sample_rows)):
        is_blank = all(
            ws.cell(row_idx, col).value is None
            for col in range(data_region["start_col"], data_region["end_col"] + 1)
        )
        if is_blank:
            blank_rows.append(row_idx)

    return blank_rows
```

**Accuracy:** 95%+ with 1000 sample rows
**Speedup:** 50-100x for large files

---

### Strategy 7: **Use openpyxl's read_only + values_only**

**Concept:** Minimize openpyxl overhead

```python
wb = openpyxl.load_workbook(
    file_path,
    read_only=True,      # ← 3x faster memory usage
    data_only=True,      # ← Skip formula objects
    keep_vba=False,      # ← Skip VBA
    keep_links=False,    # ← Skip external links
)

# Use values_only for iteration
for row in ws.iter_rows(values_only=True, max_row=10000):
    # Much faster than accessing cell objects
    pass
```

**Speedup:** 2-3x faster, 5x less memory

---

### Strategy 8: **Bypass Structure Analysis for Large Files**

**Concept:** Allow users to skip auto-detection for large files

```python
class SheetConfig:
    auto_detect: bool = True
    fast_mode: bool = False  # ← New option
    max_rows_for_analysis: int = 10_000  # ← Configurable limit

class StructureAnalyzer:
    def analyze(self, file_path, sheet, force=False, config=None):
        ws = wb[sheet]

        # Check size threshold
        if config and config.fast_mode:
            # Skip analysis, use defaults
            return self._fast_defaults(ws)

        if ws.max_row > (config.max_rows_for_analysis if config else 10_000):
            # Auto-enable fast mode for very large files
            return self._fast_defaults(ws)

        # Normal analysis
        return self._analyze_full(ws)
```

---

## Recommended Implementation

### Phase 1: Quick Wins (1-2 hours)

**1. Add row limit to _detect_data_region():**

```python
def _detect_data_region(self, ws, max_rows=10_000) -> dict:
    """Find data boundaries (limited row scan)."""
    min_row, max_row = None, None
    min_col, max_col = None, None

    # Use ws.max_row as hint
    estimated_max = ws.max_row
    scan_limit = min(max_rows, estimated_max)

    # Scan limited rows
    for row_idx, row in enumerate(ws.iter_rows(max_row=scan_limit), start=1):
        row_has_data = False
        for col_idx, cell in enumerate(row, start=1):
            if cell.value is not None:
                row_has_data = True
                if min_row is None:
                    min_row = row_idx
                max_row = row_idx
                if min_col is None or col_idx < min_col:
                    min_col = col_idx
                if max_col is None or col_idx > max_col:
                    max_col = col_idx

        # Early exit if we've seen 100 empty rows
        if min_row and row_idx - max_row > 100:
            break

    # Trust ws.max_row for the end
    if max_row and max_row < estimated_max:
        max_row = estimated_max

    return {...}
```

**Speedup:** 10-100x for large files

**2. Add sampling to _detect_blank_rows():**

```python
def _detect_blank_rows(self, ws, data_region, max_sample=1000) -> list[int]:
    """Find blank rows (sampled for large files)."""
    total_rows = data_region["end_row"] - data_region["start_row"] + 1

    if total_rows <= max_sample:
        # Small file: check all rows
        return self._detect_blank_rows_full(ws, data_region)

    # Large file: sample strategically
    # - First 300 rows (likely to have metadata/headers)
    # - Last 300 rows (likely to have totals)
    # - Random 400 from middle
    sample_rows = self._select_sample_rows(
        data_region["start_row"],
        data_region["end_row"],
        max_sample
    )

    blank_rows = []
    for row_idx in sample_rows:
        is_blank = all(
            ws.cell(row_idx, col).value is None
            for col in range(data_region["start_col"], data_region["end_col"] + 1)
        )
        if is_blank:
            blank_rows.append(row_idx)

    return sorted(blank_rows)
```

**Speedup:** 50x for 50,000 row files

**3. Use values_only iteration:**

```python
# Current (slow):
for row in ws.iter_rows():
    for cell in row:
        value = cell.value

# Optimized (3x faster):
for row in ws.iter_rows(values_only=True, max_row=10000):
    for value in row:
        # Direct value access, no cell object overhead
```

---

### Phase 2: Advanced Optimizations (4-6 hours)

**4. Parallel Sheet Processing:**

```python
from concurrent.futures import ThreadPoolExecutor

class MessyWorkbook:
    def to_dataframes(self, max_workers=4):
        """Process sheets in parallel."""
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(self._parse_sheet, name): name
                for name in self._sheet_names
            }

            results = {}
            for future in as_completed(futures):
                sheet = futures[future]
                results[sheet] = future.result()

            return results
```

**Speedup:** 3-4x for multi-sheet workbooks

**5. Add size-based strategy selector:**

```python
class AnalysisStrategy(Enum):
    FULL = "full"          # < 10,000 rows
    SAMPLED = "sampled"    # 10,000 - 100,000 rows
    MINIMAL = "minimal"    # > 100,000 rows

def select_strategy(file_size_bytes, row_count):
    if row_count < 10_000:
        return AnalysisStrategy.FULL
    elif row_count < 100_000:
        return AnalysisStrategy.SAMPLED
    else:
        return AnalysisStrategy.MINIMAL
```

**6. Cache worksheet dimensions:**

```python
class StructureAnalyzer:
    def __init__(self):
        self._dimension_cache = {}

    def _get_dimensions(self, ws):
        """Get worksheet dimensions (cached)."""
        ws_id = id(ws)
        if ws_id not in self._dimension_cache:
            self._dimension_cache[ws_id] = (ws.max_row, ws.max_column)
        return self._dimension_cache[ws_id]
```

---

### Phase 3: Nuclear Options (If still slow)

**7. Use pandas directly for very large files:**

```python
def _parse_sheet(self, sheet, config):
    # Check file size
    file_size_mb = self._file_path.stat().st_size / (1024 * 1024)

    if file_size_mb > 10:  # > 10MB
        # Skip structure analysis, use pandas directly
        return pd.read_excel(
            self._file_path,
            sheet_name=sheet,
            engine='openpyxl'
        )

    # Normal path with structure analysis
    return self._parse_with_structure(sheet, config)
```

**8. Add progress callbacks:**

```python
def analyze(self, file_path, sheet, progress_callback=None):
    """Analyze with optional progress reporting."""

    if progress_callback:
        progress_callback("Reading file dimensions...")

    ws = wb[sheet]

    if progress_callback:
        progress_callback(f"Scanning {ws.max_row:,} rows...")

    # ... analysis ...
```

---

## Recommended Immediate Fixes

### Fix 1: Add MAX_ROWS limit (5 minutes)

```python
# detection/structure_analyzer.py

class StructureAnalyzer:
    # Class constant
    MAX_ANALYSIS_ROWS = 10_000

    def _detect_data_region(self, ws) -> dict:
        """Find data boundaries with row limit."""
        # Use ws dimensions as starting point
        max_row_hint = ws.max_row
        max_col_hint = ws.max_column

        # Limit scanning
        scan_rows = min(self.MAX_ANALYSIS_ROWS, max_row_hint)

        min_row, max_row = None, None
        min_col, max_col = None, None

        for row_idx, row in enumerate(ws.iter_rows(max_row=scan_rows), start=1):
            # ... existing logic ...

        # If we hit the limit, trust ws.max_row for the end
        if scan_rows < max_row_hint:
            max_row = max_row_hint

        return {...}
```

### Fix 2: Skip blank row detection for large files (10 minutes)

```python
def _detect_multiple_tables(self, ws, data_region, header_info):
    """Detect tables (skip for large files)."""
    total_rows = data_region["end_row"] - data_region["start_row"] + 1

    if total_rows > 10_000:
        # Large file: assume single table
        return [TableInfo(
            start_row=data_region["start_row"],
            end_row=data_region["end_row"],
            start_col=data_region["start_col"],
            end_col=data_region["end_col"],
            has_header=True,
        )]

    # Normal multi-table detection
    blank_rows = self._detect_blank_rows(ws, data_region)
    # ... existing logic ...
```

---

## Performance Targets

| File Size | Rows | Current | Target | Strategy |
|-----------|------|---------|--------|----------|
| < 100KB | < 1,000 | < 1s | < 1s | Full analysis |
| 100KB - 1MB | 1K - 10K | 1-5s | < 2s | Full with limits |
| 1MB - 10MB | 10K - 100K | 10+ min | < 10s | Sampling |
| 10MB - 50MB | 100K+ | N/A | < 30s | Minimal/pandas direct |

---

## Implementation Priority

1. **Immediate (Fix today):**
   - Add MAX_ANALYSIS_ROWS = 10,000
   - Skip blank row detection for large files
   - Use values_only iteration

2. **Short-term (This week):**
   - Implement sampling strategy
   - Add fast_mode option
   - Cache dimensions

3. **Medium-term (Next week):**
   - Parallel sheet processing
   - Progressive loading
   - Chunk-based analysis

4. **Long-term (Future):**
   - Benchmark suite
   - Adaptive strategy selection
   - Memory profiling

---

## Estimated Impact

| Optimization | Effort | Speedup | File Size Impact |
|--------------|--------|---------|------------------|
| Row limits | 30 min | 10-100x | 1-10MB files |
| Sampling | 2 hours | 50x | 10-50MB files |
| values_only | 15 min | 3x | All files |
| Parallel processing | 1 hour | 3-4x | Multi-sheet |
| Fast path | 1 hour | 1000x | Simple files |

**Combined: 100-500x speedup for large files!**

Target: customers.xlsx (2.5MB) from **15+ minutes → 5-10 seconds**
