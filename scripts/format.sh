#!/usr/bin/env bash
set -euo pipefail

# Format Nim sources and examples using nimpretty.
find src examples tests -name "*.nim" -print0 | xargs -0 -n1 nimpretty --indent:2 --maxLineLen:100 --backup:off
