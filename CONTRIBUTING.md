# Contributing

Thank you for improving **Nim_AutohotkeyLearning**! This guide explains how the
codebase is organized, the conventions we follow, and how to add support for new
platform backends.

## Repository layout
- `src/core/` – shared runtime pieces (logging, scheduler), the `PlatformBackend`
  interface, and the `RuntimeContext` type.
- `src/platform/windows/` – Windows-specific implementations (backend wiring,
  WinAPI helpers for processes, input, window management, and hotkey handling).
- `src/platform/linux/` – explicit stubs that raise clear errors for unsupported
  features.
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
  logic belongs in `src/platform/windows/`; other platforms should either add an
  implementation or extend the Linux stubs with explicit errors.
- Use the shared `RuntimeContext` to thread logging, scheduling, and backend
  access through actions and plugins.
- Prefer descriptive log messages; the `Logger` from `core/logging` is the
  canonical logger for runtime code.
- Condition any Windows-only imports with `when defined(windows):` to avoid
  compile-time issues on other platforms.
- Keep modules small and focused; reuse helpers from `core/` and `features/`
  before adding new utilities.

## Adding a new platform backend
1. Create `src/platform/<platform>/backend.nim` that inherits from
   `PlatformBackend` and implements the required methods. Use
   `backendUnsupported` for operations that cannot be supported.
2. Add any platform-specific helpers (hotkeys, window management, input, etc.)
   inside `src/platform/<platform>/` and wire them through your backend
   implementation.
3. Update `src/main.nim` to instantiate the backend for the new platform (via a
   `when defined(<platform>):` branch) and ensure plugins are registered only
   where supported.
4. Keep feature modules platform-agnostic by calling `ctx.backend` instead of
   importing platform modules directly.
5. Add tests or examples if the new backend introduces new capabilities, and
   document any limitations in the relevant module headers.
