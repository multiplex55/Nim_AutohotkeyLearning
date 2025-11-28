import ../../core/platform_backend
import ../../core/scheduler

## Placeholder backend that clearly marks unsupported features on Linux.
type
  LinuxBackend* = ref object of PlatformBackend

proc newLinuxBackend*(): LinuxBackend =
  LinuxBackend()

method runMessageLoop*(backend: LinuxBackend; scheduler: Scheduler) =
  discard scheduler
  raise backendUnsupported("runMessageLoop (Linux stub)")

method registerHotkey*(backend: LinuxBackend; modifiers: int; key: int;
    cb: HotkeyCallback): HotkeyId =
  discard modifiers; discard key; discard cb
  raise backendUnsupported("registerHotkey (Linux stub)")

method postQuit*(backend: LinuxBackend; exitCode: int = 0) =
  discard exitCode
  raise backendUnsupported("postQuit (Linux stub)")

method startProcessDetached*(backend: LinuxBackend; command: string; args: seq[
    string] = @[]): bool =
  discard command; discard args
  raise backendUnsupported("startProcessDetached (Linux stub)")

method killProcessesByName*(backend: LinuxBackend; name: string;
    exitCode: int = 0): int =
  discard name; discard exitCode
  raise backendUnsupported("killProcessesByName (Linux stub)")

method sendText*(backend: LinuxBackend; text: string) =
  discard text
  raise backendUnsupported("sendText (Linux stub)")

method setMousePos*(backend: LinuxBackend; x, y: int): bool =
  discard x; discard y
  raise backendUnsupported("setMousePos (Linux stub)")

method leftClick*(backend: LinuxBackend) =
  raise backendUnsupported("leftClick (Linux stub)")

method getActiveWindow*(backend: LinuxBackend): WindowHandle =
  raise backendUnsupported("getActiveWindow (Linux stub)")

method getWindowTitle*(backend: LinuxBackend; hwnd: WindowHandle): string =
  discard hwnd
  raise backendUnsupported("getWindowTitle (Linux stub)")

method describeWindow*(backend: LinuxBackend; hwnd: WindowHandle): string =
  discard hwnd
  raise backendUnsupported("describeWindow (Linux stub)")

method centerWindowOnPrimaryMonitor*(backend: LinuxBackend;
    hwnd: WindowHandle): bool =
  discard hwnd
  raise backendUnsupported("centerWindowOnPrimaryMonitor (Linux stub)")

method getPrimaryScreenSize*(backend: LinuxBackend): tuple[width: int; height: int] =
  raise backendUnsupported("getPrimaryScreenSize (Linux stub)")
