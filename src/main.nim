import winim/lean          # VK_*, MOD_* constants, GetSystemMetrics
import hotkeys
import processes
import ./windows as win
import mouse_keyboard

const
  VK_N = 'N'.ord
  VK_K = 'K'.ord
  VK_C = 'C'.ord
  VK_T = 'T'.ord
  VK_M = 'M'.ord
  VK_L = 'L'.ord
  VK_H = 'H'.ord

when isMainModule:
  echo "Nim AHK-like demo started."
  echo "Hotkeys (if registration succeeds):"
  echo "  ESC           : exit program"
  echo "  Win+Alt+N     : start Notepad"
  echo "  Win+Alt+K     : kill all notepad.exe processes"
  echo "  Win+Alt+C     : center the active window"
  echo "  Win+Alt[T]    : print active window info"
  echo "  Win+Alt+M     : move mouse to screen center"
  echo "  Win+Alt+L     : left-click at current mouse position"
  echo "  Win+Alt+H     : type \"Hello from Nim!\""
  echo ""

  # 1) ESC = kill switch / exit
  try:
    discard registerHotkey(
      0,
      VK_ESCAPE,
      proc() =
        echo "[ESC] Kill switch pressed, exiting..."
        postQuit()
    )
    echo "Registered ESC as kill switch."
  except IOError as e:
    echo "Could not register ESC hotkey: ", e.msg

  # 2) Win+Alt+N = start Notepad
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

  # 6) Win+Alt+M = move mouse to center of primary monitor
  try:
    discard registerHotkey(
      MOD_WIN or MOD_ALT,
      VK_M,
      proc() =
        let screenW = GetSystemMetrics(SM_CXSCREEN)
        let screenH = GetSystemMetrics(SM_CYSCREEN)
        let x = screenW div 2
        let y = screenH div 2

        if setMousePos(x, y):
          echo "[Win+Alt+M] Moved mouse to screen center (", x, ", ", y, ")."
        else:
          echo "[Win+Alt+M] Failed to move mouse."
    )
    echo "Registered Win+Alt+M to move mouse to screen center."
  except IOError as e:
    echo "Could not register Win+Alt+M hotkey: ", e.msg

  # 7) Win+Alt+L = left click at current mouse position
  try:
    discard registerHotkey(
      MOD_WIN or MOD_ALT,
      VK_L,
      proc() =
        let pos = getMousePos()
        echo "[Win+Alt+L] Left click at (", pos.x, ", ", pos.y, ")."
        leftClick()
    )
    echo "Registered Win+Alt+L to left-click at current mouse position."
  except IOError as e:
    echo "Could not register Win+Alt+L hotkey: ", e.msg

  # 8) Win+Alt+H = send text "Hello from Nim!"
  try:
    discard registerHotkey(
      MOD_WIN or MOD_ALT,
      VK_H,
      proc() =
        let msg = "Hello from Nim!"
        echo "[Win+Alt+H] Sending text: \"", msg, "\""
        sendText(msg)
    )
    echo "Registered Win+Alt+H to send \"Hello from Nim!\"."
  except IOError as e:
    echo "Could not register Win+Alt+H hotkey: ", e.msg

  echo ""
  echo "Entering message loop. Close this window or use any working hotkeys."
  runMessageLoop()
  echo "Message loop exited. Goodbye."
