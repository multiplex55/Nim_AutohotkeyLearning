import std/[options, strutils, tables, times]

import ./uia
import ../actions
import ../plugins
import ../../core/[logging, runtime_context, window_targets]

when defined(windows):
  import winim/lean
  import ../../platform/windows/processes as winProcesses

## Plugin that owns the lifetime of a UI Automation session.
type
  UiaPlugin* = ref object of Plugin
    uia*: Uia

when defined(windows):
  type
    ResolvedTarget = object
      hwnd: HWND
      element: ptr IUIAutomationElement

proc newUiaPlugin*(): UiaPlugin =
  UiaPlugin(name: "uia", description: "Windows UI Automation helpers")

proc parseDurationMs(params: Table[string, string], key: string, defaultMs: int): Duration =
  if key in params:
    try:
      return initDuration(milliseconds = parseInt(params[key]))
    except ValueError:
      discard
  initDuration(milliseconds = defaultMs)

when defined(windows):
  proc parseControlTypeId(value: string): int =
    if value.len == 0:
      return UIA_ButtonControlTypeId
    let normalized = value.toLowerAscii()
    case normalized
    of "button": UIA_ButtonControlTypeId
    of "menuitem", "menu_item": UIA_MenuItemControlTypeId
    of "tabitem", "tab": UIA_TabItemControlTypeId
    else:
      try:
        parseInt(value)
      except ValueError:
        UIA_ButtonControlTypeId

  proc readWindowTitle(hwnd: HWND): string =
    var buffer = newStringOfCap(256)
    buffer.setLen(256)
    let copied = GetWindowText(hwnd, buffer.cstring, 256)
    if copied <= 0:
      return ""
    buffer.setLen(copied)
    buffer

  proc readClassName(hwnd: HWND): string =
    var buffer = newStringOfCap(256)
    buffer.setLen(256)
    let copied = GetClassName(hwnd, buffer.cstring, 256)
    if copied <= 0:
      return ""
    buffer.setLen(copied)
    buffer

  proc buildProcessMap(): Table[int, string] =
    result = initTable[int, string]()
    for p in winProcesses.enumProcesses():
      result[p.pid] = p.exeName.toLowerAscii()

  proc processNameForPid(pid: DWORD, cache: var Table[int, string]): Option[string] =
    if cache.len == 0:
      cache = buildProcessMap()
    if int(pid) in cache:
      return some(cache[int(pid)])
    none(string)

  proc hasSelectors(target: WindowTarget): bool =
    target.title.isSome or target.titleContains.isSome or
    target.className.isSome or target.processName.isSome

  proc matchesTarget(hwnd: HWND, target: WindowTarget, procMap: var Table[int, string]): bool =
    if IsWindow(hwnd) == 0:
      return false

    let title = readWindowTitle(hwnd)
    if target.title.isSome and title != target.title.get():
      return false

    if target.titleContains.isSome:
      let needle = target.titleContains.get().toLowerAscii()
      if not title.toLowerAscii().contains(needle):
        return false

    if target.className.isSome:
      let cls = readClassName(hwnd).toLowerAscii()
      if cls != target.className.get().toLowerAscii():
        return false

    if target.processName.isSome:
      var pid: DWORD
      discard GetWindowThreadProcessId(hwnd, addr pid)
      let procNameOpt = processNameForPid(pid, procMap)
      if procNameOpt.isNone or procNameOpt.get() != target.processName.get().toLowerAscii():
        return false

    result = target.hasSelectors()

  proc findWindowBySelectors(target: WindowTarget, logger: Logger): Option[HWND] =
    if not target.hasSelectors():
      if logger != nil:
        logger.warn("Window target missing selectors; cannot resolve", [("name", target.name)])
      return none(HWND)

    var matched: HWND = 0
    var procMap = initTable[int, string]()

    proc enumProc(hwnd: HWND, l: LPARAM): WINBOOL {.stdcall.} =
      discard l
      if IsWindowVisible(hwnd) == 0:
        return 1
      if matchesTarget(hwnd, target, procMap):
        matched = hwnd
        return 0
      1

    discard EnumWindows(enumProc, 0)

    if matched != 0:
      return some(matched)

    if logger != nil:
      logger.warn("No window matched target selectors", [("name", target.name)])
    none(HWND)

  proc resolveTargetWindow(plugin: UiaPlugin, targetName: string, ctx: RuntimeContext): Option[HWND] =
    if targetName.len == 0:
      if ctx.logger != nil:
        ctx.logger.warn("UIA action missing target name")
      return none(HWND)

    if targetName notin ctx.windowTargets:
      if ctx.logger != nil:
        ctx.logger.warn("Unknown window target", [("target", targetName)])
      return none(HWND)

    let target = ctx.windowTargets[targetName]

    if target.storedHwnd.isSome:
      let hwnd = HWND(target.storedHwnd.get())
      if IsWindow(hwnd) != 0:
        return some(hwnd)
      elif ctx.logger != nil:
        ctx.logger.warn("Stored HWND for target is invalid", [("target", targetName), ("hwnd", $target.storedHwnd.get())])

    let matched = findWindowBySelectors(target, ctx.logger)
    if matched.isSome and ctx.logger != nil:
      ctx.logger.info("Resolved window via selectors", [("target", targetName), ("hwnd", $cast[int](matched.get()))])
    matched

  proc resolveTargetElement(plugin: UiaPlugin, targetName: string, ctx: RuntimeContext): Option[ResolvedTarget] =
    let hwndOpt = plugin.resolveTargetWindow(targetName, ctx)
    if hwndOpt.isNone:
      return none(ResolvedTarget)

    try:
      let element = plugin.uia.fromWindowHandle(hwndOpt.get())
      return some(ResolvedTarget(hwnd: hwndOpt.get(), element: element))
    except CatchableError as e:
      if ctx.logger != nil:
        ctx.logger.error("Failed to convert HWND to UIA element", [("target", targetName), ("error", e.msg)])
      none(ResolvedTarget)

proc registerAliases(registry: var ActionRegistry, names: openArray[string], factory: ActionFactory) =
  for name in names:
    registry.registerAction(name, factory)

method install*(plugin: UiaPlugin, registry: var ActionRegistry, ctx: var RuntimeContext) =
  plugin.uia = initUia()
  if ctx.logger != nil:
    ctx.logger.info("Initialized UIA session", [("name", plugin.name)])

  when defined(windows):
    registerAliases(registry, ["uia_click_button", "uia: click button"], proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
      let targetName = params.getOrDefault("target", "").strip()
      let buttonName = params.getOrDefault("button", params.getOrDefault("name", "")).strip()
      let automationId = params.getOrDefault("automation_id", "").strip()
      let timeout = parseDurationMs(params, "timeout_ms", 3000)
      let controlTypeId = parseControlTypeId(params.getOrDefault("control_type", "Button"))

      return proc() =
        try:
          let resolvedOpt = plugin.resolveTargetElement(targetName, ctx)
          if resolvedOpt.isNone:
            return
          let resolved = resolvedOpt.get()

          var cond: ptr IUIAutomationCondition
          if automationId.len > 0:
            cond = plugin.uia.automationIdAndControlType(automationId, controlTypeId)
          elif buttonName.len > 0:
            cond = plugin.uia.nameAndControlType(buttonName, controlTypeId)
          else:
            cond = plugin.uia.controlTypeCondition(controlTypeId)

          let button = plugin.uia.waitElement(tsDescendants, cond, timeout = timeout, root = resolved.element)
          if button.isNil:
            if ctx.logger != nil:
              ctx.logger.warn("UIA button not found", [("target", targetName), ("name", buttonName), ("automationId", automationId)])
            return

          plugin.uia.invoke(button)
          if ctx.logger != nil:
            ctx.logger.info("Invoked UIA button", [("target", targetName), ("name", buttonName), ("automationId", automationId)])
        except CatchableError as e:
          if ctx.logger != nil:
            ctx.logger.error("UIA click button failed", [("target", targetName), ("error", e.msg)])
    )

    registerAliases(registry, ["uia_close_window", "uia: close window"], proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
      let targetName = params.getOrDefault("target", "").strip()
      return proc() =
        try:
          let resolvedOpt = plugin.resolveTargetElement(targetName, ctx)
          if resolvedOpt.isNone:
            return
          let resolved = resolvedOpt.get()
          resolved.element.closeWindow()
          if ctx.logger != nil:
            ctx.logger.info("Closed window via UIA", [("target", targetName), ("hwnd", $cast[int](resolved.hwnd))])
        except CatchableError as e:
          if ctx.logger != nil:
            ctx.logger.error("UIA close window failed", [("target", targetName), ("error", e.msg)])
    )

    registerAliases(registry, ["uia_query_state", "uia: query state"], proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
      let targetName = params.getOrDefault("target", "").strip()
      return proc() =
        try:
          let resolvedOpt = plugin.resolveTargetElement(targetName, ctx)
          if resolvedOpt.isNone:
            return
          let resolved = resolvedOpt.get()

          let nameVal = resolved.element.currentName()
          let classVal = resolved.element.currentClassName()
          let automationVal = resolved.element.currentAutomationId()
          let enabledVal = resolved.element.isEnabled()
          let offscreenVal = resolved.element.isOffscreen()
          let stateVal = resolved.element.windowVisualState()

          if ctx.logger != nil:
            ctx.logger.info(
              "UIA window state",
              [
                ("target", targetName),
                ("hwnd", $cast[int](resolved.hwnd)),
                ("name", nameVal),
                ("class", classVal),
                ("automationId", automationVal),
                ("enabled", $enabledVal),
                ("offscreen", $offscreenVal),
                ("visualState", $int(stateVal))
              ]
            )
        except CatchableError as e:
          if ctx.logger != nil:
            ctx.logger.error("UIA query state failed", [("target", targetName), ("error", e.msg)])
    )

    registerAliases(registry, ["uia_diagnostics", "uia: diagnostics"], proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
      let targetName = params.getOrDefault("target", "").strip()
      let timeout = parseDurationMs(params, "timeout_ms", 1000)

      return proc() =
        try:
          let resolvedOpt = plugin.resolveTargetElement(targetName, ctx)
          if resolvedOpt.isNone:
            return
          let resolved = resolvedOpt.get()

          let descriptor = ctx.backend.describeWindow(WindowHandle(resolved.hwnd))
          let nameVal = resolved.element.currentName()
          let classVal = resolved.element.currentClassName()
          let automationVal = resolved.element.currentAutomationId()

          if ctx.logger != nil:
            ctx.logger.info(
              "UIA diagnostics",
              [
                ("target", targetName),
                ("hwnd", $cast[int](resolved.hwnd)),
                ("window", descriptor),
                ("name", nameVal),
                ("class", classVal),
                ("automationId", automationVal),
                ("timeoutMs", $timeout.inMilliseconds)
              ]
            )
        except CatchableError as e:
          if ctx.logger != nil:
            ctx.logger.error("UIA diagnostics failed", [("target", targetName), ("error", e.msg)])
    )

    registerAliases(registry, ["uia_dump_element", "uia: dump element"], proc(params: Table[string, string], ctx: var RuntimeContext): TaskAction =
      let sourceParam = params.getOrDefault("source", "mouse").strip().toLowerAscii()

      return proc() =
        try:
          var source = if sourceParam.len == 0: "mouse" else: sourceParam
          var element: ptr IUIAutomationElement
          var hwnd: HWND = 0
          var procCache = initTable[int, string]()
          var mousePos: Option[tuple[x: int32, y: int32]] = none(tuple[x: int32, y: int32])

          case source
          of "active", "window":
            source = "active"
            let active = ctx.backend.getActiveWindow()
            if active == 0:
              if ctx.logger != nil:
                ctx.logger.warn("No active window to dump UIA element")
              return
            hwnd = HWND(active)
            element = plugin.uia.fromWindowHandle(hwnd)
          else:
            source = "mouse"
            var pt: POINT
            if GetCursorPos(addr pt) == 0:
              if ctx.logger != nil:
                ctx.logger.warn("Failed to read mouse position for UIA dump")
              return
            mousePos = some((x: pt.x, y: pt.y))
            element = plugin.uia.fromPoint(pt.x, pt.y)
            let nativeHandle = element.nativeWindowHandle()
            if nativeHandle != 0:
              hwnd = HWND(nativeHandle)

          let nameVal = element.currentName()
          let classVal = element.currentClassName()
          let automationVal = element.currentAutomationId()
          let controlTypeId = element.currentControlType()
          let patterns = element.availablePatterns()
          let enabledVal = element.isEnabled()
          let offscreenVal = element.isOffscreen()
          let visibleVal = element.isVisible()
          let focusableVal = element.isKeyboardFocusable()
          let focusVal = element.hasKeyboardFocus()
          let isControlVal = element.isControlElement()
          let isContentVal = element.isContentElement()
          let passwordVal = element.isPassword()

          var fields: seq[(string, string)] = @[
            ("source", source),
            ("name", nameVal),
            ("automationId", automationVal),
            ("class", classVal),
            ("controlType", controlTypeName(controlTypeId)),
            ("controlTypeId", $controlTypeId),
            ("patterns", patterns.join(",")),
            ("enabled", $enabledVal),
            ("visible", $visibleVal),
            ("offscreen", $offscreenVal),
            ("focusable", $focusableVal),
            ("hasFocus", $focusVal),
            ("isControlElement", $isControlVal),
            ("isContentElement", $isContentVal),
            ("isPassword", $passwordVal)
          ]

          if mousePos.isSome:
            let pt = mousePos.get()
            fields.add(("mouseX", $pt.x))
            fields.add(("mouseY", $pt.y))

          if hwnd != 0:
            fields.add(("hwnd", $cast[int](hwnd)))
            fields.add(("windowTitle", readWindowTitle(hwnd)))
            fields.add(("windowClass", readClassName(hwnd)))

            var pid: DWORD
            discard GetWindowThreadProcessId(hwnd, addr pid)
            let procOpt = processNameForPid(pid, procCache)
            if procOpt.isSome:
              fields.add(("processName", procOpt.get()))

          if ctx.logger != nil:
            ctx.logger.info("UIA element dump", fields)
        except CatchableError as e:
          if ctx.logger != nil:
            ctx.logger.error("UIA dump failed", [("error", e.msg)])
    )

method shutdown*(plugin: UiaPlugin, ctx: RuntimeContext) =
  discard ctx
  if plugin.uia != nil:
    plugin.uia.shutdown()
    plugin.uia = nil
