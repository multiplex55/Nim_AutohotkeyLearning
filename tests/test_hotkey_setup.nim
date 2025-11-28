import std/[options, tables, unittest]

import ../src/main
import ../src/core/[logging, runtime_context, scheduler, window_targets]
import ../src/features/actions
import ../src/core/platform_backend


type
  RecordingBackend = ref object of PlatformBackend
    registered: seq[(int, int)]
    clears: int

method registerHotkey*(backend: RecordingBackend; modifiers: int; key: int;
    cb: HotkeyCallback): HotkeyId =
  discard cb
  backend.registered.add((modifiers, key))
  HotkeyId(backend.registered.len)

method clearHotkeys*(backend: RecordingBackend) =
  inc backend.clears
  backend.registered.setLen(0)


template newHotkey(keys: string; enabled = true): HotkeyConfig =
  HotkeyConfig(
    enabled: enabled,
    keys: keys,
    action: "left_click",
    params: initTable[string, string](),
    target: "",
    uiaAction: "",
    uiaParams: initTable[string, string](),
    delayMs: none(int),
    repeatMs: none(int),
    sequence: @[]
  )

suite "hotkey setup":
  test "disabled hotkeys are not registered":
    let backend = RecordingBackend(registered: @[], clears: 0)
    let logger = newLogger()
    let sched = newScheduler(logger)
    var runtime = RuntimeContext(
      logger: logger,
      scheduler: sched,
      backend: backend,
      windowTargets: initTable[string, WindowTarget](),
      windowTargetStatePath: none(string)
    )
    var registry = newActionRegistry(logger)
    registerBuiltinActions(registry)

    var cfg = ConfigResult()
    cfg.hotkeys = @[newHotkey("Ctrl+K", enabled = true), newHotkey("Ctrl+L",
        enabled = false)]

    let registered = registerConfiguredHotkeys(cfg, backend, registry, runtime)

    check registered == 1
    check backend.registered.len == 1
    check backend.clears == 1
