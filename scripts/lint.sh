#!/usr/bin/env bash
set -euo pipefail

# Static checks for the library + examples.
nim check src/main.nim
nim check src/nim_ahkTesting.nim
