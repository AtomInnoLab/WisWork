#!/usr/bin/env bash
# Read-only local integrity audit for the 20 fixed PPT benchmark candidates.
set -euo pipefail

repo_root="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
materials="$repo_root/docs/product/ppt-benchmark-materials"
manifest_count=0
validator_count=0

cd "$repo_root"
for number in $(seq -w 1 20); do
  case_dir="$materials/PPT-P0-$number"
  test -s "$case_dir/README.md" || { echo "PPT-P0-$number: README missing" >&2; exit 1; }

  if test -f "$case_dir/SHA256SUMS"; then
    (cd "$case_dir" && sha256sum --check --status SHA256SUMS) || {
      echo "PPT-P0-$number: frozen file hash mismatch" >&2
      exit 1
    }
    manifest_count=$((manifest_count + 1))
  fi

  if test -f "$case_dir/verify-materials.mjs"; then
    node --import tsx "$case_dir/verify-materials.mjs" >/dev/null || {
      echo "PPT-P0-$number: material validator failed" >&2
      exit 1
    }
    validator_count=$((validator_count + 1))
  elif test "$number" = 11; then
    python3 "$case_dir/verify-scope-narrowed-candidate.py" >/dev/null || {
      echo "PPT-P0-11: scope validator failed" >&2
      exit 1
    }
    validator_count=$((validator_count + 1))
  fi
  echo "PPT-P0-$number: local material checks passed"
done

echo "Summary: 20 candidate directories, $manifest_count hash manifests, $validator_count case validators passed. Professional review and PowerPoint host acceptance remain separate."
