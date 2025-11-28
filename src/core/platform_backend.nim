import std/strformat

import ./scheduler

## Base error type for platform-related failures.
type
  PlatformError* = object of CatchableError

  ## Cross-platform window handle alias.
  WindowHandle* = int

  ## Type for platform-specific backends.
  PlatformBackend* = ref object of RootObj

  ## Hotkey callback primitives shared across backends.
  HotkeyId* = int32
  HotkeyCallback* = proc() {.closure.}

proc backendUnsupported*(feature: string): ref PlatformError =
  ## Helper to raise a consistent unsupported-platform error.
  newException(PlatformError, fmt"Feature '{feature}' is not supported on this platform (hostOS={hostOS}).")

method startProcessDetached*(backend: PlatformBackend; command: string;
    args: seq[string] = @[]): bool {.base.} =
  raise backendUnsupported("startProcessDetached")

method killProcessesByName*(backend: PlatformBackend; name: string;
    exitCode: int = 0): int {.base.} =
  raise backendUnsupported("killProcessesByName")

method sendText*(backend: PlatformBackend; text: string) {.base.} =
  raise backendUnsupported("sendText")

method setMousePos*(backend: PlatformBackend; x, y: int): bool {.base.} =
  raise backendUnsupported("setMousePos")

method leftClick*(backend: PlatformBackend) {.base.} =
  raise backendUnsupported("leftClick")

method getActiveWindow*(backend: PlatformBackend): WindowHandle {.base.} =
  raise backendUnsupported("getActiveWindow")

method getWindowTitle*(backend: PlatformBackend;
    hwnd: WindowHandle): string {.base.} =
  raise backendUnsupported("getWindowTitle")

method describeWindow*(backend: PlatformBackend;
    hwnd: WindowHandle): string {.base.} =
  raise backendUnsupported("describeWindow")

method centerWindowOnPrimaryMonitor*(backend: PlatformBackend;
    hwnd: WindowHandle): bool {.base.} =
  raise backendUnsupported("centerWindowOnPrimaryMonitor")

method getPrimaryScreenSize*(backend: PlatformBackend): tuple[width: int;
    height: int] {.base.} =
  raise backendUnsupported("getPrimaryScreenSize")

method registerHotkey*(backend: PlatformBackend; modifiers: int; key: int;
    cb: HotkeyCallback): HotkeyId {.base.} =
  raise backendUnsupported("registerHotkey")

method clearHotkeys*(backend: PlatformBackend) {.base.} =
  discard

method runMessageLoop*(backend: PlatformBackend;
    scheduler: Scheduler) {.base.} =
  raise backendUnsupported("runMessageLoop")

method postQuit*(backend: PlatformBackend; exitCode: int = 0) {.base.} =
  raise backendUnsupported("postQuit")
