# Package

version = "0.1.0"
author = "multiplex55"
description = "ahk like NIM"
license = "MIT"
srcDir = "src"
bin = @["nim_ahkTesting"]


# Dependencies

requires "nim >= 2.2.6"
requires "parsetoml >= 0.7.0"
requires "winim >= 3.9.3"

task lint, "Run static analysis":
  exec "nim check src/main.nim"

task test, "Run unit tests":
  exec "nim c -r tests/test_win_integration.nim"
  exec "nim c -r tests/test_window_handles.nim"
  exec "nim c -r tests/test_uia.nim"
