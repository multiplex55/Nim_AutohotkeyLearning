import std/[unittest]

when defined(windows):
  import win/win
  import platform/windows/windows as nativeWindows
  import platform/windows/processes

  suite "win handle safety":
    test "invalid handles return safe defaults":
      let invalidHwnd = WindowHandle(0x1234)
      check windowInfo(invalidHwnd).isNone
      check not activateWindow(invalidHwnd)
      check not moveResizeWindow(invalidHwnd, 0, 0, 10, 10)

    test "null handles are ignored safely":
      check windowInfo(WindowHandle(0)).isNone
      check not activateWindow(WindowHandle(0))
      check not moveResizeWindow(WindowHandle(0), 0, 0, 5, 5)

  suite "native window helper guardrails":
    test "invalid native handles are handled":
      let invalidHwnd = nativeWindows.WindowHandle(0x1234)
      check nativeWindows.getWindowTitle(invalidHwnd) == ""
      let rect = nativeWindows.getWindowRect(invalidHwnd)
      check rect.width == 0 and rect.height == 0
      check not nativeWindows.moveWindow(invalidHwnd, 0, 0, 50, 50)
      check not nativeWindows.centerWindowOnPrimaryMonitor(invalidHwnd)
      check not nativeWindows.bringToFront(invalidHwnd)
      check "invalid window" in nativeWindows.describeWindow(invalidHwnd)

  suite "process helper guardrails":
    test "killProcessByPid gracefully rejects invalid pid":
      check not killProcessByPid(-1)
else:
  static:
    echo "Skipping handle safety tests on non-Windows platforms."
