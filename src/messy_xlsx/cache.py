"""LRU cache for structure analysis results."""

# ============================================================================
# Imports
# ============================================================================

from __future__ import annotations

from collections import OrderedDict
from dataclasses import dataclass
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


@dataclass(frozen=True)
class PathIdentity:
    """Stable filesystem identity captured around structure analysis."""

    resolved: str
    device: int
    inode: int
    size: int
    mtime_ns: int
    ctime_ns: int

    @classmethod
    def before(cls, path: Path) -> PathIdentity:
        """Capture all path and stat fields used by the global cache."""
        resolved_path = path.resolve()
        stat = resolved_path.stat()
        return cls(
            resolved=str(resolved_path),
            device=stat.st_dev,
            inode=stat.st_ino,
            size=stat.st_size,
            mtime_ns=stat.st_mtime_ns,
            ctime_ns=stat.st_ctime_ns,
        )

    def unchanged(self, path: Path) -> bool:
        """Return whether the path still has this exact identity."""
        try:
            return self == type(self).before(path)
        except OSError:
            return False


class StructureCache:
    """Specialized cache for StructureInfo results."""

    def __init__(self, maxsize: int = 128):
        self._cache: LRUCache[StructureInfo] = LRUCache(maxsize)

    def _make_key(
        self,
        file_path: Path,
        sheet: str,
        identity: PathIdentity | None = None,
        variant: str | None = None,
    ) -> str:
        """Create a key from exact path identity, sheet, and analysis variant."""
        if identity is None:
            identity = PathIdentity.before(file_path)
        stat_key = (
            identity.device,
            identity.inode,
            identity.size,
            identity.mtime_ns,
            identity.ctime_ns,
        )
        return f"{identity.resolved}:{sheet}:{stat_key}:{variant or ''}"

    def get(
        self,
        file_path: Path,
        sheet: str,
        variant: str | None = None,
    ) -> StructureInfo | None:
        """Get cached structure info for a sheet."""
        try:
            identity = PathIdentity.before(file_path)
        except OSError:
            return None

        key = self._make_key(file_path, sheet, identity, variant)
        cached = self._cache.get(key)
        if cached is None or not identity.unchanged(file_path):
            return None
        return cached

    def put(
        self,
        file_path: Path,
        sheet: str,
        info: StructureInfo,
        variant: str | None = None,
        identity: PathIdentity | None = None,
    ) -> bool:
        """Cache structure info for a sheet."""
        try:
            before = identity or PathIdentity.before(file_path)
        except OSError:
            return False

        if not before.unchanged(file_path):
            return False

        key = self._make_key(file_path, sheet, before, variant)
        self._cache.put(key, info)
        return True

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
