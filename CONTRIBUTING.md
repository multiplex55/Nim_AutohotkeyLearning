# Contributing

Thank you for improving **Nim_AutohotkeyLearning**! This guide explains how the
codebase is organized and the conventions we follow for the Windows-only
tooling.

## Repository layout
- `src/core/` – shared runtime pieces (logging, scheduler), the `PlatformBackend`
  interface, and the `RuntimeContext` type.
- `src/platform/windows/` – Windows-specific implementations (backend wiring,
  WinAPI helpers for processes, input, window management, and hotkey handling).
- `src/features/` – feature modules that are mostly platform-agnostic
  (action registry, config parsing, key parsing, plugins) plus feature-focused
  subfolders:
  - `features/input/` – higher-level input helpers.
  - `features/uia/` – UIA helpers (currently Windows-only).
  - `features/win_automation/` – Windows automation plugins built on the
    platform backend.
- `src/examples/` – example configs and sample programs.

## Coding standards
- Keep platform-specific code behind the `PlatformBackend` interface. Windows
  logic belongs in `src/platform/windows/`.
- Use the shared `RuntimeContext` to thread logging, scheduling, and backend
  access through actions and plugins.
- Prefer descriptive log messages; the `Logger` from `core/logging` is the
  canonical logger for runtime code.
- Keep modules small and focused; reuse helpers from `core/` and `features/`
  before adding new utilities.

## Platform scope

The project intentionally targets **Windows 10/11 (x64)** only. If you plan to
experiment with additional platforms, please discuss in an issue first; out of
the box there are no non-Windows backends or stubs.
