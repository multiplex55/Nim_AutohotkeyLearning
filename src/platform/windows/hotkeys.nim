# hotkeys.nim
## Simple global hotkey helper for Windows (Nim).
# hotkeys.nim
## Simple global hotkey helper for Windows (Nim).
##
## Features:
## - Register global hotkeys (modifier + key)
## - Map each hotkey to a Nim callback
## - Run a blocking Windows message loop that dispatches WM_HOTKEY
##
## This is Windows-only and requires the `winim` package.
##
## Basic usage:
##   import hotkeys
##   import winim/lean  # for VK_* and MOD_* constants
##
##   discard registerHotkey(0, VK_ESCAPE, proc() =
##     echo "ESC pressed, exiting..."
##     postQuit()
##   )
##
##   runMessageLoop()   # blocks until postQuit() is called

import std/tables
import winim/lean
import ../../core/scheduler
import ../../core/logging
from ../../core/platform_backend import HotkeyCallback, HotkeyId

var
  hotkeyCallbacks: Table[HotkeyId, HotkeyCallback]
  nextId: HotkeyId = 1'i32
  runningLoop = false

proc pollHotkeyMessages*(scheduler: Scheduler = nil) =
  ## Non-blocking hotkey dispatcher that can be called from custom
  ## message loops (e.g. GUI frameworks) to service WM_HOTKEY events.
  ##
  ## Unlike `runMessageLoop`, this does not block and only drains
  ## WM_HOTKEY messages from the current thread's queue.
  var msg: MSG

  while PeekMessage(addr msg, HWND(0), WM_HOTKEY, WM_HOTKEY, PM_REMOVE) != 0:
    let id = HotkeyId(msg.wParam)

    if scheduler != nil and scheduler.logger != nil:
      scheduler.logger.debug("WM_HOTKEY received", [("id", $id)])

    if id in hotkeyCallbacks:
      let cb = hotkeyCallbacks[id]
      if cb != nil:
        cb()

proc registerHotkey*(modifiers: int, vk: int, cb: HotkeyCallback): HotkeyId =
  ## Register a new global hotkey.
  ##
  ## - `modifiers` is a bitwise OR of MOD_* flags,
  ##    e.g. `MOD_CONTROL or MOD_ALT`. It can also be 0 for a bare key.
  ## - `vk` is the virtual-key code (e.g. VK_ESCAPE, 0x51 for 'Q').
  ## - `cb` is the callback invoked whenever the hotkey is pressed.
  ##
  ## Returns:
  ##   A HotkeyId which you can later pass to `unregisterHotkey`.
  ##
  ## Raises:
  ##   IOError if Windows rejects the registration (e.g. another program
  ##   already uses that hotkey).
  ##
  ## Note:
  ##   You must call this *before* `runMessageLoop()` starts.
  if runningLoop:
    raise newException(IOError, "Cannot register new hotkeys after runMessageLoop() has started")

  let id = nextId
  inc nextId

  # WinAPI: proc RegisterHotKey(hWnd: HWND; id: int32; fsModifiers: UINT; vk: UINT): WINBOOL
  let ok = RegisterHotKey(HWND(0), id, UINT(modifiers), UINT(vk))
  if ok == 0:
    raise newException(IOError, "RegisterHotKey failed (hotkey may already be in use)")

  hotkeyCallbacks[id] = cb
  result = id

proc unregisterHotkey*(id: HotkeyId) =
  ## Unregister a previously registered global hotkey.
  ##
  ## Safe to call even if `id` does not exist.
  if id in hotkeyCallbacks:
    discard UnregisterHotKey(HWND(0), id)
    hotkeyCallbacks.del(id)

proc unregisterAllHotkeys*() =
  ## Unregister all registered global hotkeys.
  ##
  ## Call this if you want to clean up before exiting, though Windows
  ## will clean these up when the process terminates.
  for id in hotkeyCallbacks.keys:
    discard UnregisterHotKey(HWND(0), id)
  hotkeyCallbacks.clear()

proc runMessageLoop*(scheduler: Scheduler = nil) =
  ## Start a simple Windows message loop that dispatches WM_HOTKEY messages.
  ##
  ## This call blocks until `postQuit()` (or `PostQuitMessage`) is called.
  ##
  ## Typical flow:
  ##   1. Register your hotkeys and callbacks.
  ##   2. Register an exit hotkey that calls `postQuit()`, e.g. ESC.
  ##   3. Call `runMessageLoop()` to start listening.
  ##
  ## You should only call this once in a given program.
  var msg: MSG

  runningLoop = true
  defer:
    runningLoop = false
    unregisterAllHotkeys()

  while true:
    # Non-blocking message pump so the scheduler can make progress.
    if PeekMessage(addr msg, HWND(0), 0, 0, PM_REMOVE) != 0:
      if msg.message == WM_QUIT:
        break

      if msg.message == WM_HOTKEY:
        let id = HotkeyId(msg.wParam)

        if scheduler != nil and scheduler.logger != nil:
          scheduler.logger.debug("WM_HOTKEY received", [("id", $id)])

        if id in hotkeyCallbacks:
          let cb = hotkeyCallbacks[id]
          if cb != nil:
            cb()

      TranslateMessage(addr msg)
      DispatchMessage(addr msg)
    else:
      if scheduler != nil:
        scheduler.tick()
      Sleep(1)

proc postQuit*(exitCode: int = 0) =
  ## Signal the message loop started by `runMessageLoop()` to exit.
  ##
  ## This is a thin wrapper around `PostQuitMessage`.
  ##
  ## You typically call this from within a hotkey callback, e.g.
  ##
  ##   registerHotkey(0, VK_ESCAPE, proc() =
  ##     echo "ESC pressed"
  ##     postQuit()
  ##   )
  PostQuitMessage(int32(exitCode))
