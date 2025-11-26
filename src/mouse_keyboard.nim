## mouse_keyboard.nim
## AHK-like mouse and keyboard helpers for Windows.
##
## Features:
##   - Query and set mouse position
##   - Left/right click (at current position or at a specific point)
##   - Basic key down/up / key press
##   - Simple ASCII text sending (letters, digits, space, newline)
##
## This is intentionally small and focused. It uses the older
## `mouse_event` and `keybd_event` APIs which are easy to call
## and perfectly fine for automation purposes.

when system.hostOS != "windows":
  {.error: "mouse_keyboard.nim only supports Windows.".}

import winim/lean

type
  ## Simple struct to hold mouse coordinates in screen space.
  MousePoint* = object
    x*: int
    y*: int

# ─────────────────────────────────────────────────────────────────────────────
# Raw WinAPI wrappers (mouse_event / keybd_event)
# ─────────────────────────────────────────────────────────────────────────────

## Wrapper around the Win32 mouse_event function.
## We give it a Nim-friendly name and signature.
proc mouseEvent(dwFlags, dx, dy, dwData: DWORD; dwExtraInfo: ULONG_PTR)
  {.stdcall, dynlib: "user32", importc: "mouse_event".}

## Wrapper around the Win32 keybd_event function.
proc keybdEvent(bVk: BYTE; bScan: BYTE; dwFlags: DWORD; dwExtraInfo: ULONG_PTR)
  {.stdcall, dynlib: "user32", importc: "keybd_event".}

## Mouse event flags (subset).
const
  MOUSEEVENTF_LEFTDOWN*  = DWORD(0x0002)
  MOUSEEVENTF_LEFTUP*    = DWORD(0x0004)
  MOUSEEVENTF_RIGHTDOWN* = DWORD(0x0008)
  MOUSEEVENTF_RIGHTUP*   = DWORD(0x0010)

## Keyboard event flags (subset).
const
  KEYEVENTF_KEYUP* = DWORD(0x0002)

# ─────────────────────────────────────────────────────────────────────────────
# Common virtual-key constants (exported for convenience)
# ─────────────────────────────────────────────────────────────────────────────
##
## These are meant to be convenient, AHK-style key names you can use
## anywhere you import `mouse_keyboard`. They are just ints underneath,
## so they work directly with `registerHotkey`, `sendKeyPress`, etc.
##
## Example usage:
##   import mouse_keyboard, hotkeys
##   discard registerHotkey(MOD_WIN or MOD_ALT, KEY_N, proc() = echo "Win+Alt+N")
##
## You can still use the raw VK_* constants from winim if you prefer.

const
  # No modifier / placeholder
  KEY_NONE* = 0

  # Letters A–Z
  KEY_A* = 'A'.ord
  KEY_B* = 'B'.ord
  KEY_C* = 'C'.ord
  KEY_D* = 'D'.ord
  KEY_E* = 'E'.ord
  KEY_F* = 'F'.ord
  KEY_G* = 'G'.ord
  KEY_H* = 'H'.ord
  KEY_I* = 'I'.ord
  KEY_J* = 'J'.ord
  KEY_K* = 'K'.ord
  KEY_L* = 'L'.ord
  KEY_M* = 'M'.ord
  KEY_N* = 'N'.ord
  KEY_O* = 'O'.ord
  KEY_P* = 'P'.ord
  KEY_Q* = 'Q'.ord
  KEY_R* = 'R'.ord
  KEY_S* = 'S'.ord
  KEY_T* = 'T'.ord
  KEY_U* = 'U'.ord
  KEY_V* = 'V'.ord
  KEY_W* = 'W'.ord
  KEY_X* = 'X'.ord
  KEY_Y* = 'Y'.ord
  KEY_Z* = 'Z'.ord

  # Digits 0–9
  KEY_0* = '0'.ord
  KEY_1* = '1'.ord
  KEY_2* = '2'.ord
  KEY_3* = '3'.ord
  KEY_4* = '4'.ord
  KEY_5* = '5'.ord
  KEY_6* = '6'.ord
  KEY_7* = '7'.ord
  KEY_8* = '8'.ord
  KEY_9* = '9'.ord

  # Control / editing keys
  KEY_ESCAPE*    = VK_ESCAPE
  KEY_ESC*       = VK_ESCAPE
  KEY_ENTER*     = VK_RETURN
  KEY_RETURN*    = VK_RETURN
  KEY_SPACE*     = VK_SPACE
  KEY_TAB*       = VK_TAB
  KEY_BACKSPACE* = VK_BACK
  KEY_DELETE*    = VK_DELETE
  KEY_INSERT*    = VK_INSERT

  # Modifiers
  KEY_SHIFT*     = VK_SHIFT
  KEY_CONTROL*   = VK_CONTROL
  KEY_CTRL*      = VK_CONTROL
  KEY_ALT*       = VK_MENU

  # Navigation / arrows
  KEY_LEFT*      = VK_LEFT
  KEY_RIGHT*     = VK_RIGHT
  KEY_UP*        = VK_UP
  KEY_DOWN*      = VK_DOWN
  KEY_HOME*      = VK_HOME
  KEY_END*       = VK_END
  KEY_PAGEUP*    = VK_PRIOR
  KEY_PAGEDOWN*  = VK_NEXT

  # Function keys
  KEY_F1*  = VK_F1
  KEY_F2*  = VK_F2
  KEY_F3*  = VK_F3
  KEY_F4*  = VK_F4
  KEY_F5*  = VK_F5
  KEY_F6*  = VK_F6
  KEY_F7*  = VK_F7
  KEY_F8*  = VK_F8
  KEY_F9*  = VK_F9
  KEY_F10* = VK_F10
  KEY_F11* = VK_F11
  KEY_F12* = VK_F12

# ─────────────────────────────────────────────────────────────────────────────
# Mouse helpers
# ─────────────────────────────────────────────────────────────────────────────

proc getMousePos*(): MousePoint =

  ## Returns the current mouse position in screen coordinates.
  ##
  ## If GetCursorPos fails (rare), returns (0, 0).
  var pt: POINT
  if GetCursorPos(addr pt) != 0:
    result = MousePoint(x: int(pt.x), y: int(pt.y))
  else:
    result = MousePoint(x: 0, y: 0)


proc setMousePos*(x, y: int): bool =
  ## Moves the mouse cursor to the absolute screen coordinates (x, y).
  ##
  ## Returns true on success, false otherwise.
  result = SetCursorPos(int32(x), int32(y)) != 0


proc leftClick*() =
  ## Performs a left mouse click at the current cursor position.
  mouseEvent(
    MOUSEEVENTF_LEFTDOWN or MOUSEEVENTF_LEFTUP,
    DWORD(0), DWORD(0), DWORD(0), ULONG_PTR(0)
  )

proc rightClick*() =
  ## Performs a right mouse click at the current cursor position.
  mouseEvent(
    MOUSEEVENTF_RIGHTDOWN or MOUSEEVENTF_RIGHTUP,
    DWORD(0), DWORD(0), DWORD(0), ULONG_PTR(0)
  )

proc leftClickAt*(x, y: int) =
  ## Moves the cursor to (x, y) and performs a left click.
  discard setMousePos(x, y)
  leftClick()

proc rightClickAt*(x, y: int) =
  ## Moves the cursor to (x, y) and performs a right click.
  discard setMousePos(x, y)
  rightClick()


# ─────────────────────────────────────────────────────────────────────────────
# Keyboard helpers
# ─────────────────────────────────────────────────────────────────────────────

proc keyDown*(vk: int) =
  ## Simulate key down for the given virtual-key code.
  ##
  ## Example:
  ##   keyDown(VK_SHIFT)
  keybdEvent(BYTE(vk), BYTE(0), DWORD(0), ULONG_PTR(0))

proc keyUp*(vk: int) =
  ## Simulate key up for the given virtual-key code.
  ##
  ## Example:
  ##   keyUp(VK_SHIFT)
  keybdEvent(BYTE(vk), BYTE(0), KEYEVENTF_KEYUP, ULONG_PTR(0))

proc sendKeyPress*(vk: int) =
  ## Send a full key press (down + up) for the given virtual-key code.
  ##
  ## Example:
  ##   sendKeyPress(VK_RETURN)
  keyDown(vk)
  keyUp(vk)


proc sendText*(text: string) =
  ## Very simple text sender for US-layout-ish keyboards.
  ##
  ## Supports:
  ##   - 'a'..'z'   -> sent as lowercase letters
  ##   - 'A'..'Z'   -> sent as uppercase letters (Shift + key)
  ##   - '0'..'9'   -> sent as digits
  ##   - ' '        -> space
  ##   - '\n'       -> Enter
  ##
  ## Any other characters are currently ignored.

  # Helper: map a letter (a-z or A-Z) to its virtual-key code (VK_A..VK_Z).
  proc letterVk(ch: char): int =
    ## On Windows, VK codes for letters are always the uppercase code.
    ## 'a'..'z' map to 'A'..'Z' by subtracting 32 in ASCII.
    if ch in 'a'..'z':
      result = ord(ch) - 32     # 'a'(97) -> 'A'(65)
    elif ch in 'A'..'Z':
      result = ord(ch)
    else:
      result = ord(ch)

  for ch in text:
    case ch
    of 'a'..'z':
      # Lowercase letter: press the key without Shift.
      let vk = letterVk(ch)
      sendKeyPress(vk)

    of 'A'..'Z':
      # Uppercase letter: hold Shift while pressing the key.
      let vk = letterVk(ch)
      keyDown(VK_SHIFT)
      sendKeyPress(vk)
      keyUp(VK_SHIFT)

    of '0'..'9':
      # Digits map directly to their VK_* codes on US layout.
      sendKeyPress(ord(ch))

    of ' ':
      sendKeyPress(VK_SPACE)

    of '\n':
      sendKeyPress(VK_RETURN)

    else:
      # Unsupported character – ignore for now.
      discard

