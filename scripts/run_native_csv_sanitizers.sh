#!/usr/bin/env bash
set -euo pipefail

if [[ "${1:-}" != "--shell-only" || "$#" -ne 1 ]]; then
  echo "usage: $0 --shell-only" >&2
  exit 2
fi

script_dir="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd -P)"
project_root="$(cd "$script_dir/.." && pwd -P)"
python="$project_root/.venv/bin/python"
scratch="$(mktemp -d "${TMPDIR:-/tmp}/messy-xlsx-native-sanitizers-XXXXXX")"
trap 'rm -rf "$scratch"' EXIT

MESSY_XLSX_BUILD_MODE=fallback "$python" -m build \
  --sdist \
  --outdir "$scratch/dist" \
  "$project_root"
mkdir -p "$scratch/source"
tar -xzf "$scratch"/dist/*.tar.gz -C "$scratch/source"
source_tree="$(find "$scratch/source" -mindepth 1 -maxdepth 1 -type d)"
"$python" -m venv "$scratch/venv"

compiler="${CC:-cc}"
asan_input="${MESSY_XLSX_ASAN_LIBRARY:-$("$compiler" -print-file-name=libasan.so)}"
ubsan_input="${MESSY_XLSX_UBSAN_LIBRARY:-$("$compiler" -print-file-name=libubsan.so)}"
sanitizer_library_dir="${MESSY_XLSX_SANITIZER_LIBRARY_DIR:-}"

resolve_runtime_library() {
  local compiler_input="$1"
  local linked_runtime
  if file -L "$compiler_input" | grep -q "ELF .* shared object"; then
    printf '%s\n' "$compiler_input"
    return 0
  fi
  linked_runtime="$(sed -n 's/.*( *\\([^ )]*\\).*/\\1/p' "$compiler_input")"
  if [[ -f "$linked_runtime" ]] \
    && file -L "$linked_runtime" | grep -q "ELF .* shared object"; then
    printf '%s\n' "$linked_runtime"
    return 0
  fi
  return 1
}

if ! asan_library="$(resolve_runtime_library "$asan_input")" \
  || ! ubsan_library="$(resolve_runtime_library "$ubsan_input")"; then
  echo "ASan/UBSan runtime libraries are unavailable for $compiler" >&2
  exit 1
fi

sanitizer_flags="-O1 -g -fno-omit-frame-pointer -fsanitize=address,undefined"
sanitizer_link_flags="-fsanitize=address,undefined"
if [[ -n "$sanitizer_library_dir" ]]; then
  sanitizer_link_flags="-L$sanitizer_library_dir -Wl,-rpath,$sanitizer_library_dir $sanitizer_link_flags"
fi
CFLAGS="$sanitizer_flags" \
LDFLAGS="$sanitizer_link_flags" \
MESSY_XLSX_BUILD_MODE=native \
uv pip install \
  --python "$scratch/venv/bin/python" \
  -e "$source_tree[dev]" \
  -r "$source_tree/requirements/native-release.txt"

LD_PRELOAD="$asan_library:$ubsan_library" \
LD_LIBRARY_PATH="$sanitizer_library_dir${LD_LIBRARY_PATH:+:$LD_LIBRARY_PATH}" \
ASAN_OPTIONS="detect_leaks=0:halt_on_error=1:allocator_may_return_null=1" \
UBSAN_OPTIONS="halt_on_error=1:print_stacktrace=1" \
PYTHONMALLOC=debug \
"$scratch/venv/bin/python" -m pytest \
  "$source_tree/tests/native_csv/test_abi_shell.py" \
  -q
