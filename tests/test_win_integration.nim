import std/[unittest]

when defined(windows):
  import win/win
  import input/input

  suite "win helpers":
    test "find and wait returns none safely":
      check findWindowByTitle("nonexistent").isNone
      check winWait("also-not-there", 100.milliseconds).isNone

    test "list windows returns seq":
      discard listWindows()

  suite "input helpers":
    test "keyboard and mouse stubs":
      # These calls should not raise even if they don't target a real window.
      moveMouse(MousePoint(x: 0, y: 0))
      clickMouse()
      dragMouse(MousePoint(x: 0, y: 0), MousePoint(x: 10, y: 10), steps = 1)
      scrollMouse(0)
      sendKeys([VK_SHIFT])
      typeText("hi")
else:
  static: echo "Skipping Windows integration tests on non-Windows platforms."
