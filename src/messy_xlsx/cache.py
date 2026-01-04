"""LRU cache for structure analysis results."""

# ============================================================================
# Imports
# ============================================================================

from collections import OrderedDict
from pathlib import Path
from threading import RLock
from typing import Generic, TypeVar

from messy_xlsx.models import StructureInfo


# ============================================================================
# Type Variables
# ============================================================================

T = TypeVar("T")


# ============================================================================
# Generic LRU Cache
# ============================================================================

class LRUCache(Generic[T]):
    """Thread-safe LRU (Least Recently Used) cache."""

    def __init__(self, maxsize: int = 128):
        self.maxsize = maxsize
        self._cache: OrderedDict[str, T] = OrderedDict()
        self._lock = RLock()

    def get(self, key: str) -> T | None:
        """Get value from cache."""
        with self._lock:
            if key not in self._cache:
                return None
            self._cache.move_to_end(key)
            return self._cache[key]

    def put(self, key: str, value: T) -> None:
        """Add or update value in cache."""
        with self._lock:
            if key in self._cache:
                self._cache.move_to_end(key)
                self._cache[key] = value
            else:
                if len(self._cache) >= self.maxsize:
                    self._cache.popitem(last=False)
                self._cache[key] = value

    def invalidate(self, key: str) -> bool:
        """Remove a specific key from cache."""
        with self._lock:
            if key in self._cache:
                del self._cache[key]
                return True
            return False

    def invalidate_prefix(self, prefix: str) -> int:
        """Remove all keys starting with given prefix."""
        with self._lock:
            keys_to_remove = [k for k in self._cache if k.startswith(prefix)]
            for key in keys_to_remove:
                del self._cache[key]
            return len(keys_to_remove)

    def clear(self) -> None:
        """Remove all entries from cache."""
        with self._lock:
            self._cache.clear()

    def __len__(self) -> int:
        """Return number of cached entries."""
        with self._lock:
            return len(self._cache)

    def __contains__(self, key: str) -> bool:
        """Check if key is in cache."""
        with self._lock:
            return key in self._cache


# ============================================================================
# Structure-Specific Cache
# ============================================================================

class StructureCache:
    """Specialized cache for StructureInfo results."""

    def __init__(self, maxsize: int = 128):
        self._cache: LRUCache[StructureInfo] = LRUCache(maxsize)

    def _make_key(self, file_path: Path, sheet: str, mtime: float | None = None) -> str:
        """Create cache key from file path, sheet, and mtime."""
        if mtime is None:
            try:
                mtime = file_path.stat().st_mtime
            except OSError:
                mtime = 0.0
        return f"{file_path.resolve()}:{sheet}:{mtime}"

    def get(self, file_path: Path, sheet: str) -> StructureInfo | None:
        """Get cached structure info for a sheet."""
        try:
            mtime = file_path.stat().st_mtime
        except OSError:
            return None

        key = self._make_key(file_path, sheet, mtime)
        return self._cache.get(key)

    def put(self, file_path: Path, sheet: str, info: StructureInfo) -> None:
        """Cache structure info for a sheet."""
        try:
            mtime = file_path.stat().st_mtime
        except OSError:
            return

        key = self._make_key(file_path, sheet, mtime)
        self._cache.put(key, info)

    def invalidate(self, file_path: Path) -> int:
        """Invalidate all cached entries for a file."""
        prefix = str(file_path.resolve()) + ":"
        return self._cache.invalidate_prefix(prefix)

    def clear(self) -> None:
        """Remove all cached entries."""
        self._cache.clear()

    def __len__(self) -> int:
        """Return number of cached entries."""
        return len(self._cache)


# ============================================================================
# Global Cache Instance
# ============================================================================

_structure_cache = StructureCache()


def get_structure_cache() -> StructureCache:
    """Get the global structure cache instance."""
    return _structure_cache


def clear_cache() -> None:
    """Clear the global structure cache."""
    _structure_cache.clear()
