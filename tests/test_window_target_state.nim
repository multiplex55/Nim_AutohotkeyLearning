import std/[os, tables, unittest]

import ../src/core/window_targets
import ../src/features/window_target_state

suite "window target state persistence":
  test "invalid handles are ignored when loading state":
    let tmpDir = getTempDir()
    let statePath = joinPath(tmpDir, "window_target_invalid.toml")

    writeFile(statePath, """
[window_targets.bad]
hwnd = -42
""")

    var targets = initTable[string, WindowTarget]()
    loadWindowTargetState(statePath, targets, nil)

    check "bad" notin targets

    discard tryRemoveFile(statePath)

  test "valid handles merge into the target map":
    let tmpDir = getTempDir()
    let statePath = joinPath(tmpDir, "window_target_valid.toml")

    writeFile(statePath, """
[window_targets.editor]
hwnd = 123456
""")

    var targets = initTable[string, WindowTarget]()
    loadWindowTargetState(statePath, targets, nil)

    check "editor" in targets
    check targets["editor"].storedHwnd.isSome
    check targets["editor"].storedHwnd.get() == 123456

    discard tryRemoveFile(statePath)
