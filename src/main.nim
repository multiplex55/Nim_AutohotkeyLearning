import winim/lean          # for VK_* and MOD_* constants
import hotkeys
import processes
import ./windows as win    # our Windows helper module

const
  VK_C = 'C'.ord
  VK_T = 'T'.ord

when isMainModule:
  echo "Nim AHK-like demo started."
  echo "Hotkeys (if registration succeeds):"
  echo "  ESC           : exit program"
  echo "  Win+Alt+N     : start Notepad"
  echo "  Win+Alt+K     : kill all notepad.exe processes"
  echo "  Win+Alt+C     : center the active window"
  echo "  Win+Alt+T     : print active window info"
  echo ""

  # 1) ESC = kill switch / exit
  try:
    discard registerHotkey(
      0,               # no modifiers
      VK_ESCAPE,
      proc() =
        echo "[ESC] Kill switch pressed, exiting..."
        postQuit()
    )
    echo "Registered ESC as kill switch."
  except IOError as e:
    echo "Could not register ESC hotkey: ", e.msg

  # 2) Win+Alt+N = start Notepad
  const VK_N = 'N'.ord
  try:
    discard registerHotkey(
      MOD_WIN or MOD_ALT,
      VK_N,
      proc() =
        echo "[Win+Alt+N] Starting notepad.exe..."
        if startProcessDetached("notepad.exe"):
          echo "  -> Started Notepad."
        else:
          echo "  -> Failed to start Notepad."
    )
    echo "Registered Win+Alt+N to start Notepad."
  except IOError as e:
    echo "Could not register Win+Alt+N hotkey: ", e.msg

  # 3) Win+Alt+K = kill all Notepad processes
  const VK_K = 'K'.ord
  try:
    discard registerHotkey(
      MOD_WIN or MOD_ALT,
      VK_K,
      proc() =
        echo "[Win+Alt+K] Attempting to kill notepad.exe processes..."
        let killed = killProcessesByName("notepad.exe")
        echo "  -> Killed ", killed, " processes."
    )
    echo "Registered Win+Alt+K to kill Notepad."
  except IOError as e:
    echo "Could not register Win+Alt+K hotkey: ", e.msg

  # 4) Win+Alt+C = center active window
  try:
    discard registerHotkey(
      MOD_WIN or MOD_ALT,
      VK_C,
      proc() =
        let hwnd = win.getActiveWindow()
        if hwnd == 0:
          echo "[Win+Alt+C] No active window detected."
          return

        let title = win.getWindowTitle(hwnd)
        if win.centerWindowOnPrimaryMonitor(hwnd):
          echo "[Win+Alt+C] Centered active window: \"", title, "\""
        else:
          echo "[Win+Alt+C] Failed to center window: \"", title, "\""
    )
    echo "Registered Win+Alt+C to center the active window."
  except IOError as e:
    echo "Could not register Win+Alt+C hotkey: ", e.msg

  # 5) Win+Alt+T = print active window info
  try:
    discard registerHotkey(
      MOD_WIN or MOD_ALT,
      VK_T,
      proc() =
        let hwnd = win.getActiveWindow()
        if hwnd == 0:
          echo "[Win+Alt+T] No active window detected."
        else:
          echo "[Win+Alt+T] Active window info: ", win.describeWindow(hwnd)
    )
    echo "Registered Win+Alt+T to print active window info."
  except IOError as e:
    echo "Could not register Win+Alt+T hotkey: ", e.msg

  echo ""
  echo "Entering message loop. Close this window or use any working hotkeys."
  runMessageLoop()
  echo "Message loop exited. Goodbye."
