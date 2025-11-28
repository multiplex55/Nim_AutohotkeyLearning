## ahk_dsl.nim
## Lightweight Autohotkey-inspired helpers built on top of win and input modules.

import std/[options, times]

when defined(windows):
  import ../platform/windows/win
  import ./input/input

  type
    WindowSession* = object
      title*: string
      handle*: WindowHandle
      delays*: InputDelays

  proc refreshHandle(ws: var WindowSession) =
    if ws.handle == 0:
      let found = findWindowByTitle(ws.title)
      if found.isSome:
        ws.handle = found.get.handle

  proc withWindow*(title: string; delays: InputDelays = defaultDelays): WindowSession =
    var handle: WindowHandle = 0
    let found = findWindowByTitle(title)
    if found.isSome:
      handle = found.get.handle
    WindowSession(title: title, handle: handle, delays: delays)

  proc ensureActive(ws: var WindowSession): bool =
    ws.refreshHandle()
    if ws.handle == 0:
      return false
    activateWindow(ws.handle)

  proc sendKeys*(ws: var WindowSession; keys: openArray[int]) =
    if ensureActive(ws):
      hotkey(keys, ws.delays)

  proc typeText*(ws: var WindowSession; text: string) =
    if ensureActive(ws):
      input.typeText(text, ws.delays)

  proc clickMouse*(ws: var WindowSession; pos: Option[MousePoint] = none(MousePoint)) =
    if ensureActive(ws):
      input.clickMouse(pos = pos, delays = ws.delays)

  proc winWait*(title: string; timeout: Duration = 3.seconds): Option[
      WindowSession] =
    let found = win.winWait(title, timeout)
    if found.isSome:
      let info = found.get
      return some(WindowSession(title: info.title, handle: info.handle,
          delays: defaultDelays))
    none(WindowSession)
else:
  {.warning: "ahk_dsl is only available on Windows targets.".}

  type WindowSession* = object

  proc withWindow*(title: string; delays: InputDelays = defaultDelays): WindowSession =
    discard title; discard delays
    raise newException(OSError, "Autohotkey-style helpers require Windows.")

  proc winWait*(title: string; timeout: Duration = 3.seconds): Option[
      WindowSession] =
    discard title; discard timeout
    none(WindowSession)
