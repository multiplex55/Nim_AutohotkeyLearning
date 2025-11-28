# Nim AHK Toolkit – How-to & Recipes

Practical steps for installing, building, and extending the toolkit. Use this alongside the demo in `src/main.nim` to quickly iterate on AutoHotkey-style automations.

## Introduction
- Automate Windows from Nim with global hotkeys, window control, process management, and UI Automation helpers.
- Modules are intentionally small: compose them in your own binaries or copy/paste the snippets below into new files under `examples/`.

## Installation & prerequisites
1) **OS**: Windows 10/11 (64-bit).
2) **Nim**: Install Nim 2.x from [nim-lang.org](https://nim-lang.org/install.html) and ensure `nim --version` works in a new shell.
3) **winim**: Install the WinAPI bindings once via Nimble:
   ```bash
   nimble install winim
   ```
4) Optional: clone this repo and verify the nimble file resolves dependencies:
   ```bash
   git clone https://github.com/your-org/Nim_AutohotkeyLearning.git
   cd Nim_AutohotkeyLearning
   nimble refresh
   ```

## Building & running the demo
From the repo root:
```bash
cd src
nim c -d:release main.nim           # build
nim c -r -d:release main.nim         # build + run the hotkey demo
```
Hotkey reminders (Windows): `ESC` exits; `Ctrl+Alt+A` launches Notepad; `Ctrl+Alt+Q` kills all `notepad.exe`; `Ctrl+Alt+W` centers the active window.

## Configuring & running UIA demos
UI Automation support lives in `src/uia.nim` and is wired into the demo via the `--uia-demo` flag.
1) Ensure you are on Windows and the `UIAutomationCore` runtime is available (it ships with Windows).
2) Build and run the demo with UIA enabled:
   ```bash
   cd src
   nim c -r -d:release main.nim --uia-demo
   ```
3) With a target app (e.g., Notepad) in focus, use the UIA hotkeys defined in `main.nim` to invoke or inspect controls. The console prints the element name, control type, AutomationId, and native HWND.
4) If your app already initialized COM differently, call `initUia(coinit = COINIT_MULTITHREADED)` (or your chosen flag) in your own entrypoint.

### UIA capture workflow (tree + selector)
Use the bundled hotkeys in `examples/hotkeys.toml` to quickly map an app's automation surface:
1) **Bind the active window**: Press `Ctrl+Alt+B` (`capture_window_target`) to store the foreground window as your target profile (e.g., `notepad`).
2) **Dump the UIA tree**: Press `Ctrl+Alt+T` (`uia_dump_tree`) to log a structured outline of the active/target window up to four levels deep. Each node now prints a composed selector containing the runtime ID, control type, name, and AutomationId so you can identify the same element across runs.
3) **Capture a specific element**: Hover the mouse over the control you want and press `Ctrl+Alt+C` (`uia_capture_element`). The logger records a ready-to-use selector such as `runtimeId=[42,99,3], automationId="Save", name="Save", controlType=Button`. Use `Ctrl+Alt+Shift+C` to run the same capture and copy the selector directly to the Windows clipboard for pasting into scripts or notes.

#### Using the runtime ID selector in config
The selector string is composable: it always includes `runtimeId=[...]` plus any available AutomationId and name. Practical ways to use it:
- **Repeatable targeting for `invoke`**: paste the captured selector into your notes, then transfer the AutomationId/control type pieces into the `uia_params` block:
  ```toml
  [[hotkeys]]
  name = "Invoke Save"
  keys = "Ctrl+Alt+U"
  uia_action = "invoke"
  target = "notepad"
  uia_params.automation_id = "Save"    # from selector
  uia_params.control_type = "Button"   # from selector
  ```
- **Quick recall from tree dumps**: every `uia_dump_tree` line shows `selector=runtimeId=[...], automationId=..., name=..., controlType=...`. Copy the selector text to the clipboard (or re-run `uia_capture_element` with `copy_selector = true`) and keep it alongside your automation config for reliable cross-run matching.

## Module walk-throughs
Each snippet is runnable on Windows; paste into a new `.nim` file and execute with `nim c -r filename.nim`.

### 1) Registering hotkeys (`src/hotkeys.nim`)
```nim
import hotkeys
import winim/lean      # MOD_* constants
import mouse_keyboard  # KEY_* constants

# Register Ctrl+Alt+R to print a message
if registerHotkey(MOD_CONTROL or MOD_ALT, KEY_R, proc() = echo "Ctrl+Alt+R pressed!"):
  echo "Hotkey registered. Press ESC to exit."
else:
  quit "Hotkey already in use"

# Exit when ESC is pressed
discard registerHotkey(0, KEY_ESCAPE, proc() = postQuit())
runMessageLoop()
```
Expected output after pressing the hotkey:
```
Hotkey registered. Press ESC to exit.
Ctrl+Alt+R pressed!
```

### 2) Process enumeration & termination (`src/processes.nim`)
```nim
import processes

# List processes
for p in enumProcesses():
  echo p.pid, " -> ", p.exeName

# Find + terminate Notepad (case-insensitive)
let matches = findProcessesByName("notepad.exe")
if matches.len == 0:
  echo "No Notepad found"
else:
  let killed = killProcessesByName("notepad.exe")
  echo "Killed ", killed, " instance(s)"
```
⚠️ **Safety**: `killProcessesByName` terminates *all* matching processes; double-check the name before running.

### 3) Window search, activation, and centering (`src/windows.nim`)
```nim
import windows as win

let hwnd = win.getActiveWindow()
if hwnd == 0: quit "No active window"

# Describe and center the active window
echo win.describeWindow(hwnd)
if win.centerWindowOnPrimaryMonitor(hwnd):
  echo "Centered!"

# Search by exact title and activate
let target = win.findWindowByTitleExact("Untitled - Notepad")
if target != 0 and win.bringToFront(target):
  echo "Notepad activated"
```
Expected description example:
```
HWND=0x0002019C, title="Untitled - Notepad", x=200, y=120, w=800, h=600
Centered!
Notepad activated
```

### 4) Mouse & keyboard actions (`src/mouse_keyboard.nim`)
```nim
import mouse_keyboard

# Move and click
discard setMousePos(600, 400)
leftClick()

# Send keystrokes
sendKeyPress(KEY_F5)
keyDown(KEY_SHIFT)
sendKeyPress(KEY_A)
keyUp(KEY_SHIFT)

# Simple ASCII text send (US keyboard only)
sendText("Hello from Nim!\n123")
```
⚠️ `sendText` is ASCII-only and assumes a US keyboard layout. For Unicode/IME text, use application-specific APIs instead.

### 5) UIA element interactions (`src/uia.nim`)
```nim
import std/times
import uia

when defined(windows):
  let automation = initUia()   # STA by default
  defer: automation.shutdown()

  # Wait for a button named "OK" and click it within 3 seconds
  let okButton = automation.waitElement(tsDescendants, automation.nameAndControlType("OK", UIA_ButtonControlTypeId), 3.seconds)
  if okButton != nil:
    okButton.invoke()

  # Find a search box by AutomationId and type into it
  if let elem = automation.findFirstByAutomationId("SearchEdit", tsDescendants):
    elem.setValue("Hello from Nim")
else:
  echo "UI Automation only works on Windows."
```

## Practical scenarios
### Bind a hotkey to center the active window and type text
```nim
import hotkeys, windows as win, mouse_keyboard, winim/lean

discard registerHotkey(MOD_CONTROL or MOD_ALT, KEY_C, proc() =
  let hwnd = win.getActiveWindow()
  if hwnd != 0 and win.centerWindowOnPrimaryMonitor(hwnd):
    echo "Centered active window"
    sendText("Window centered!\n")
)

discard registerHotkey(0, KEY_ESCAPE, proc() = postQuit())
runMessageLoop()
```

### Capture and reuse window targets
```nim
import windows as win

let current = win.getActiveWindow()
if current != 0:
  echo "Captured HWND: ", current
  # Later: bring it back to the foreground
  discard win.bringToFront(current)
```

### Simple automation script template
```nim
# save as examples/auto_demo.nim
import hotkeys, processes, windows as win, mouse_keyboard, winim/lean

discard registerHotkey(MOD_CONTROL or MOD_ALT, KEY_N, proc() =
  if startProcessDetached("notepad.exe"):
    echo "Launched Notepad"
)

discard registerHotkey(MOD_CONTROL or MOD_ALT, KEY_W, proc() =
  let hwnd = win.findWindowByTitleExact("Untitled - Notepad")
  if hwnd != 0:
    discard win.bringToFront(hwnd)
    sendText("Automated hello!\n")
)

discard registerHotkey(0, KEY_ESCAPE, proc() = postQuit())
runMessageLoop()
```
Run it with:
```bash
nim c -r -d:release examples/auto_demo.nim
```
Safety notes: avoid overlapping system hotkeys (Windows uses many Win+* combos). Prefer `Ctrl+Alt+<key>` to reduce conflicts.

## Troubleshooting
- **Hotkey registration fails**: another app already registered the combination. Pick a different key chord or close the conflicting app.
- **Processes not found**: ensure the executable name matches the running process (e.g., `notepad.exe`, not `Notepad`).
- **Window activation no-op**: some apps block `SetForegroundWindow`; try Alt+Tabbing once to grant focus permission or simulate `Alt` keypress before activation.
- **UIA errors**: HRESULT `RPC_E_CHANGED_MODE` means COM was initialized with a different apartment model. Re-run `initUia` using the same `coinit` flag as the host application.
- **Text send issues**: `sendText` cannot emit Unicode/IME characters. For non-ASCII input, call app-specific APIs or use UIA `setValue` where supported.
- **Antivirus prompts**: automation APIs can look suspicious; whitelist your built binary if needed.
