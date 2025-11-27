when system.hostOS != "windows":
  {.error: "Windows backend can only be built on Windows targets.".}

import winim/lean

import ../../core/platform_backend
import ../../core/scheduler
import ./hotkeys
import ./mouse_keyboard as winInput
import ./processes as winProcesses
import ./windows as winWindows except WindowHandle

## Windows implementation of PlatformBackend.
type
  WindowsBackend* = ref object of PlatformBackend

proc newWindowsBackend*(): WindowsBackend =
  WindowsBackend()

method startProcessDetached*(backend: WindowsBackend; command: string; args: seq[string] = @[]): bool =
  discard backend
  winProcesses.startProcessDetached(command, args)

method killProcessesByName*(backend: WindowsBackend; name: string; exitCode: int = 0): int =
  discard backend
  winProcesses.killProcessesByName(name, exitCode)

method sendText*(backend: WindowsBackend; text: string) =
  discard backend
  winInput.sendText(text)

method setMousePos*(backend: WindowsBackend; x, y: int): bool =
  discard backend
  winInput.setMousePos(x, y)

method leftClick*(backend: WindowsBackend) =
  discard backend
  winInput.leftClick()

method getActiveWindow*(backend: WindowsBackend): WindowHandle =
  discard backend
  WindowHandle(winWindows.getActiveWindow())

method getWindowTitle*(backend: WindowsBackend; hwnd: WindowHandle): string =
  discard backend
  winWindows.getWindowTitle(HWND(hwnd))

method describeWindow*(backend: WindowsBackend; hwnd: WindowHandle): string =
  discard backend
  winWindows.describeWindow(HWND(hwnd))

method centerWindowOnPrimaryMonitor*(backend: WindowsBackend; hwnd: WindowHandle): bool =
  discard backend
  winWindows.centerWindowOnPrimaryMonitor(HWND(hwnd))

method getPrimaryScreenSize*(backend: WindowsBackend): tuple[width: int, height: int] =
  discard backend
  (width: GetSystemMetrics(SM_CXSCREEN).int, height: GetSystemMetrics(SM_CYSCREEN).int)

method registerHotkey*(backend: WindowsBackend; modifiers: int; key: int; cb: HotkeyCallback): HotkeyId =
  discard backend
  hotkeys.registerHotkey(modifiers, key, cb)

method runMessageLoop*(backend: WindowsBackend; scheduler: Scheduler) =
  discard backend
  hotkeys.runMessageLoop(scheduler)

method postQuit*(backend: WindowsBackend; exitCode: int = 0) =
  discard backend
  hotkeys.postQuit(exitCode)
