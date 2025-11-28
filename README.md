# ⚡ Nim AHK Toolkit – AutoHotkey-style Windows automation in Nim

> ⌨️ Global hotkeys • 🪟 Window control • 🧠 Process management • 🖱️ Mouse & keyboard simulation

A small Nim library + demo app that brings AutoHotkey-style automation to Windows using pure Nim and WinAPI. For deeper guides, module references, and troubleshooting, head to **[howto.md](howto.md)**.

## 🚀 Quick start

**Requirements**
- Windows 10/11 (64-bit)
- Nim 2.x
- [`winim`](https://github.com/khchen/winim)

```bash
nimble install winim
```

**Build & run** (from repo root)

```bash
cd src
nim c -d:release main.nim         # build
nim c -r -d:release main.nim       # build + run
nim c -r -d:release main.nim --uia-demo  # UI Automation tree demo (Windows only)
```

## ✨ Features at a glance
- Register global hotkeys mapped to Nim callbacks
- Enumerate, start, and terminate processes
- Discover, move, and activate windows
- Drive mouse/keyboard input via SendInput
- UI Automation helpers for control discovery and patterns

## 🎮 Demo hotkeys (default build)
- `ESC`: Exit
- `Ctrl+Alt+A`: Launch Notepad
- `Ctrl+Alt+Q`: Kill all `notepad.exe`
- `Ctrl+Alt+W`: Center active window
- `Ctrl+Alt+E`: Print active window info
- `Ctrl+Alt+M`: Move mouse to screen center
- `Ctrl+Alt+L`: Left-click at cursor
- `Ctrl+Alt+T`: Type "Hello from Nim!"

## 📚 More
For architecture diagrams, detailed module references, UIA prerequisites, benchmarks, and extended examples, see **[howto.md](howto.md)**.
