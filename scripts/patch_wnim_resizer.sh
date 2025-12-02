#!/usr/bin/env bash
set -euo pipefail

# Patch wNim's wResizer.nim to avoid Nim 2.x inferring HashSet[None].
# The script locates Nimble-installed wNim packages and rewrites the
# initHashSet call to specify the wResizable element type explicitly.

root_dir="${HOME}/.nimble/pkgs2"
patched_any=0
shopt -s nullglob
for dir in "$root_dir"/wnim-*/wNim/private; do
  file="$dir/wResizer.nim"
  if [[ -f "$file" ]]; then
    if grep -q "initHashSet(64)" "$file"; then
      sed -i 's/initHashSet(64)/initHashSet[wResizable](64)/' "$file"
      echo "Patched: $file"
      patched_any=1
    else
      echo "No initHashSet(64) call found in $file"
    fi
  fi
done

if [[ $patched_any -eq 0 ]]; then
  echo "No wResizer.nim files patched. Ensure wNim is installed under $root_dir." >&2
fi
