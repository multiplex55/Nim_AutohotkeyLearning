# main.nim
## Test harness for hotkeys.nim + processes.nim
##
## Hotkeys:
##   - ESC           : exit program (global kill switch, if registration works)
##   - Win+Alt+N     : start Notepad
##   - Win+Alt+K     : kill all notepad.exe processes
##
## If any hotkey fails to register (e.g. already in use), we log it and continue.

import winim/lean        # for VK_* and MOD_* constants
import hotkeys
import processes

when isMainModule:
  echo "Nim AHK-like demo started."
  echo "Hotkeys (if registration succeeds):"
  echo "  ESC           : exit program"
  echo "  Win+Alt+N     : start Notepad"
  echo "  Win+Alt+K     : kill all notepad.exe processes"
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
  const VK_N = 0x4E  # 'N'
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
  const VK_K = 0x4B  # 'K'
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

  echo ""
  echo "Entering message loop. Close this window or use any working hotkeys."
  runMessageLoop()
  echo "Message loop exited. Goodbye."
