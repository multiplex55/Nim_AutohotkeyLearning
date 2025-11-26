## processes.nim
## Simple AHK-like process helpers for Windows using WinAPI + Nim stdlib.
##
## Features:
##   - Enumerate all running processes
##   - Find processes by executable name
##   - Kill processes by PID or name
##   - Start a process "detached" (no console window, best-effort)
##
## Dependencies:
##   - Nim stdlib: strutils, osproc
##   - winim: Windows API bindings (nimble install winim)

import std/[strutils, osproc]
import winim               ## <-- full winim, not winim/lean

type
  ## Basic information about a running process.
  ProcessInfo* = object
    pid*: int              ## OS process id
    exeName*: string       ## Executable filename (e.g. "notepad.exe")

# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

## Convert a WCHAR[] buffer (Windows wide-char string) to a Nim string.
proc wideToString(chars: openArray[WCHAR]): string =
  result = ""
  for c in chars:
    if cast[char](c) == '\0':
      break
    result.add(cast[char](c))

# ─────────────────────────────────────────────────────────────────────────────
# Enumeration
# ─────────────────────────────────────────────────────────────────────────────

## Enumerate all running processes using Toolhelp32 snapshot API.
proc enumProcesses*(): seq[ProcessInfo] =
  result = @[]

  var entry: PROCESSENTRY32
  var snapshot: HANDLE

  entry.dwSize = cast[DWORD](sizeof(PROCESSENTRY32))
  snapshot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)

  if snapshot == INVALID_HANDLE_VALUE:
    return

  defer: CloseHandle(snapshot)

  ## Process32First returns a WINBOOL (0/1), which Nim happily treats as bool.
  if Process32First(snapshot, addr entry):
    while true:
      let name = wideToString(entry.szExeFile)
      result.add ProcessInfo(
        pid: int(entry.th32ProcessID),
        exeName: name
      )

      if not Process32Next(snapshot, addr entry):
        break

# ─────────────────────────────────────────────────────────────────────────────
# Queries
# ─────────────────────────────────────────────────────────────────────────────

## Find all processes whose executable name matches `exeName` (case-insensitive).
##
## Example:
##   let notepads = findProcessesByName("notepad.exe")
proc findProcessesByName*(exeName: string): seq[ProcessInfo] =
  let target = exeName.toLower
  for p in enumProcesses():
    if p.exeName.toLower == target:
      result.add p

# ─────────────────────────────────────────────────────────────────────────────
# Killing processes
# ─────────────────────────────────────────────────────────────────────────────

## Kill a process by PID. Returns true on success.
proc killProcessByPid*(pid: int; exitCode: int = 0): bool =
  let hProc = OpenProcess(PROCESS_TERMINATE, FALSE, cast[DWORD](pid))
  if hProc == 0:
    return false

  defer: CloseHandle(hProc)

  ## TerminateProcess returns non-zero on success
  result = bool(TerminateProcess(hProc, cast[UINT](exitCode)))

## Kill all processes whose executable name matches `exeName`.
## Returns how many processes were successfully terminated.
##
## Example:
##   let killed = killProcessesByName("notepad.exe")
proc killProcessesByName*(exeName: string; exitCode: int = 0): int =
  result = 0
  for p in findProcessesByName(exeName):
    if killProcessByPid(p.pid, exitCode):
      inc result

# ─────────────────────────────────────────────────────────────────────────────
# Spawning processes
# ─────────────────────────────────────────────────────────────────────────────

## Start a process "detached-ish":
##   - Uses PATH (like typing in Win+R)
##   - Uses poDaemon so it doesn't pop a console window.
## Returns true if startProcess didn't throw.
##
## NOTE: We don't wait for the child at all; it just runs independently.
proc startProcessDetached*(command: string; args: openArray[string] = []): bool =
  try:
    discard startProcess(
      command,
      args = args,
      options = {poUsePath, poDaemon}
    )
    result = true
  except OSError:
    result = false
