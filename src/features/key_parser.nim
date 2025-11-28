import std/[strutils, tables]
import winim/lean
import ../platform/windows/mouse_keyboard as winKeys

const
  KEY_ESCAPE = winKeys.KEY_ESCAPE
  KEY_SPACE = winKeys.KEY_SPACE
  KEY_TAB = winKeys.KEY_TAB
  KEY_ENTER = winKeys.KEY_ENTER
  KEY_RETURN = winKeys.KEY_RETURN
  KEY_UP = winKeys.KEY_UP
  KEY_DOWN = winKeys.KEY_DOWN
  KEY_LEFT = winKeys.KEY_LEFT
  KEY_RIGHT = winKeys.KEY_RIGHT
  KEY_F1 = winKeys.KEY_F1
  KEY_F2 = winKeys.KEY_F2
  KEY_F3 = winKeys.KEY_F3
  KEY_F4 = winKeys.KEY_F4
  KEY_F5 = winKeys.KEY_F5
  KEY_F6 = winKeys.KEY_F6
  KEY_F7 = winKeys.KEY_F7
  KEY_F8 = winKeys.KEY_F8
  KEY_F9 = winKeys.KEY_F9
  KEY_F10 = winKeys.KEY_F10
  KEY_F11 = winKeys.KEY_F11
  KEY_F12 = winKeys.KEY_F12

type
  ParsedHotkey* = object
    modifiers*: int
    key*: int

let modifierMap = {
  "ctrl": MOD_CONTROL,
  "control": MOD_CONTROL,
  "alt": MOD_ALT,
  "shift": MOD_SHIFT,
  "win": MOD_WIN
}.toTable

let keyMap = block:
  var tbl = initTable[string, int]()
  for ch in 'A'..'Z':
    tbl[$ch] = ch.int
  for ch in '0'..'9':
    tbl[$ch] = ch.int
  tbl["esc"] = KEY_ESCAPE
  tbl["escape"] = KEY_ESCAPE
  tbl["space"] = KEY_SPACE
  tbl["tab"] = KEY_TAB
  tbl["enter"] = KEY_ENTER
  tbl["return"] = KEY_RETURN
  tbl["up"] = KEY_UP
  tbl["down"] = KEY_DOWN
  tbl["left"] = KEY_LEFT
  tbl["right"] = KEY_RIGHT
  tbl["f1"] = KEY_F1
  tbl["f2"] = KEY_F2
  tbl["f3"] = KEY_F3
  tbl["f4"] = KEY_F4
  tbl["f5"] = KEY_F5
  tbl["f6"] = KEY_F6
  tbl["f7"] = KEY_F7
  tbl["f8"] = KEY_F8
  tbl["f9"] = KEY_F9
  tbl["f10"] = KEY_F10
  tbl["f11"] = KEY_F11
  tbl["f12"] = KEY_F12
  tbl

proc parseHotkeyString*(raw: string): ParsedHotkey =
  ## Parse strings like "Ctrl+Alt+A" into modifiers and a virtual key.
  let parts = raw.split('+')
  var modifiers = 0
  var key = 0

  for token in parts:
    let t = token.strip.toLowerAscii()
    if t in modifierMap:
      modifiers = modifiers or modifierMap[t]
    elif t in keyMap:
      key = keyMap[t]
    elif t.len == 1:
      key = t[0].toUpperAscii.int

  result = ParsedHotkey(modifiers: modifiers, key: key)
