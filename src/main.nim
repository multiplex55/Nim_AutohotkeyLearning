import winim/lean          # VK_*, MOD_* constants, GetSystemMetrics
import hotkeys
import processes
import ./windows as win
import mouse_keyboard      # provides KEY_* constants

const
  CTRL_ALT = MOD_CONTROL or MOD_ALT

when isMainModule:
  echo "Nim AHK-like demo started."
  echo "Hotkeys (if registration succeeds):"
  echo "  ESC             : exit program"
  echo "  Ctrl+Alt+A      : start Notepad"
  echo "  Ctrl+Alt+Q      : kill all notepad.exe processes"
  echo "  Ctrl+Alt+W      : center the active window"
  echo "  Ctrl+Alt+E      : print active window info"
  echo "  Ctrl+Alt+M      : move mouse to screen center"
  echo "  Ctrl+Alt+L      : left-click at current mouse position"
  echo "  Ctrl+Alt+T      : type \"Hello from Nim!\""
  echo ""

  # 1) ESC = kill switch / exit
  try:
    discard registerHotkey(
      0,              # no modifiers
      KEY_ESCAPE,     # from mouse_keyboard (alias for VK_ESCAPE)
      proc() =
        echo "[ESC] Kill switch pressed, exiting..."
        postQuit()
    )
    echo "Registered ESC as kill switch."
  except IOError as e:
    echo "Could not register ESC hotkey: ", e.msg

  # 2) Ctrl+Alt+A = start Notepad
  try:
    discard registerHotkey(
      CTRL_ALT,
      KEY_A,
      proc() =
        echo "[Ctrl+Alt+A] Starting notepad.exe..."
        if startProcessDetached("notepad.exe"):
          echo "  -> Started Notepad."
        else:
          echo "  -> Failed to start Notepad."
    )
    echo "Registered Ctrl+Alt+A to start Notepad."
  except IOError as e:
    echo "Could not register Ctrl+Alt+A hotkey: ", e.msg

  # 3) Ctrl+Alt+Q = kill all Notepad processes
  try:
    discard registerHotkey(
      CTRL_ALT,
      KEY_Q,
      proc() =
        echo "[Ctrl+Alt+Q] Attempting to kill notepad.exe processes..."
        let killed = killProcessesByName("notepad.exe")
        echo "  -> Killed ", killed, " processes."
    )
    echo "Registered Ctrl+Alt+Q to kill Notepad."
  except IOError as e:
    echo "Could not register Ctrl+Alt+Q hotkey: ", e.msg

  # 4) Ctrl+Alt+W = center active window
  try:
    discard registerHotkey(
      CTRL_ALT,
      KEY_W,
      proc() =
        let hwnd = win.getActiveWindow()
        if hwnd == 0:
          echo "[Ctrl+Alt+W] No active window detected."
          return

        let title = win.getWindowTitle(hwnd)
        if win.centerWindowOnPrimaryMonitor(hwnd):
          echo "[Ctrl+Alt+W] Centered active window: \"", title, "\""
        else:
          echo "[Ctrl+Alt+W] Failed to center window: \"", title, "\""
    )
    echo "Registered Ctrl+Alt+W to center the active window."
  except IOError as e:
    echo "Could not register Ctrl+Alt+W hotkey: ", e.msg

  # 5) Ctrl+Alt+E = print active window info
  try:
    discard registerHotkey(
      CTRL_ALT,
      KEY_E,
      proc() =
        let hwnd = win.getActiveWindow()
        if hwnd == 0:
          echo "[Ctrl+Alt+E] No active window detected."
        else:
          echo "[Ctrl+Alt+E] Active window info: ", win.describeWindow(hwnd)
    )
    echo "Registered Ctrl+Alt+E to print active window info."
  except IOError as e:
    echo "Could not register Ctrl+Alt+E hotkey: ", e.msg

  # 6) Ctrl+Alt+M = move mouse to center of primary monitor
  try:
    discard registerHotkey(
      CTRL_ALT,
      KEY_M,
      proc() =
        let screenW = GetSystemMetrics(SM_CXSCREEN)
        let screenH = GetSystemMetrics(SM_CYSCREEN)
        let x = screenW div 2
        let y = screenH div 2

        if setMousePos(x, y):
          echo "[Ctrl+Alt+M] Moved mouse to screen center (", x, ", ", y, ")."
        else:
          echo "[Ctrl+Alt+M] Failed to move mouse."
    )
    echo "Registered Ctrl+Alt+M to move mouse to screen center."
  except IOError as e:
    echo "Could not register Ctrl+Alt+M hotkey: ", e.msg

  # 7) Ctrl+Alt+L = left click at current mouse position
  try:
    discard registerHotkey(
      CTRL_ALT,
      KEY_L,
      proc() =
        let pos = getMousePos()
        echo "[Ctrl+Alt+L] Left click at (", pos.x, ", ", pos.y, ")."
        leftClick()
    )
    echo "Registered Ctrl+Alt+L to left-click at current mouse position."
  except IOError as e:
    echo "Could not register Ctrl+Alt+L hotkey: ", e.msg

  # 8) Ctrl+Alt+T = send text "Hello from Nim!"
  try:
    discard registerHotkey(
      CTRL_ALT,
      KEY_T,
      proc() =
        let msg = "Hello from Nim!"
        echo "[Ctrl+Alt+T] Sending text: \"", msg, "\""
        sendText(msg)
    )
    echo "Registered Ctrl+Alt+T to send \"Hello from Nim!\"."
  except IOError as e:
    echo "Could not register Ctrl+Alt+T hotkey: ", e.msg

  echo ""
  echo "Entering message loop. Close this window or use any working hotkeys."
  runMessageLoop()
  echo "Message loop exited. Goodbye."
