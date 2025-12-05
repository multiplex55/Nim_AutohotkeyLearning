import std/[options, strformat, strutils, tables]

import ../actions
import ../plugins
import ../../core/runtime_context
import ../../core/platform_backend
import ../../core/scheduler
import ../../core/logging
import ../../platform/windows/mouse_keyboard as winMouse
import ./uia

import winim/lean
import winim/com
import winim/inc/uiautomation
import winim/inc/winuser

# Add Windows-specific UIA helpers here when needed.

type
  UiaPlugin* = ref object of Plugin
    uia*: Uia

proc newUiaPlugin*(): UiaPlugin =
  UiaPlugin(name: "uia", description: "Windows UI Automation helpers")

proc controlTypeFromString(value: string): int =
  let normalized = value.strip().toLowerAscii()
  case normalized
  of "button": UIA_ButtonControlTypeId
  of "edit": UIA_EditControlTypeId
  of "document": UIA_DocumentControlTypeId
  of "text": UIA_TextControlTypeId
  of "pane": UIA_PaneControlTypeId
  of "window": UIA_WindowControlTypeId
  of "tabitem": UIA_TabItemControlTypeId
  of "tab": UIA_TabControlTypeId
  of "checkbox": UIA_CheckBoxControlTypeId
  of "radiobutton": UIA_RadioButtonControlTypeId
  of "menuitem": UIA_MenuItemControlTypeId
  of "menu": UIA_MenuControlTypeId
  else:
    try:
      parseInt(normalized)
    except ValueError:
      -1

proc controlTypeName*(typeId: int): string =
  case typeId
  of UIA_ButtonControlTypeId: "Button"
  of UIA_CalendarControlTypeId: "Calendar"
  of UIA_CheckBoxControlTypeId: "CheckBox"
  of UIA_ComboBoxControlTypeId: "ComboBox"
  of UIA_DataGridControlTypeId: "DataGrid"
  of UIA_DocumentControlTypeId: "Document"
  of UIA_EditControlTypeId: "Edit"
  of UIA_GroupControlTypeId: "Group"
  of UIA_HyperlinkControlTypeId: "Hyperlink"
  of UIA_ImageControlTypeId: "Image"
  of UIA_ListControlTypeId: "List"
  of UIA_ListItemControlTypeId: "ListItem"
  of UIA_MenuControlTypeId: "Menu"
  of UIA_MenuBarControlTypeId: "MenuBar"
  of UIA_MenuItemControlTypeId: "MenuItem"
  of UIA_PaneControlTypeId: "Pane"
  of UIA_ProgressBarControlTypeId: "ProgressBar"
  of UIA_RadioButtonControlTypeId: "RadioButton"
  of UIA_ScrollBarControlTypeId: "ScrollBar"
  of UIA_SplitButtonControlTypeId: "SplitButton"
  of UIA_StatusBarControlTypeId: "StatusBar"
  of UIA_TabControlTypeId: "Tab"
  of UIA_TabItemControlTypeId: "TabItem"
  of UIA_TextControlTypeId: "Text"
  of UIA_TitleBarControlTypeId: "TitleBar"
  of UIA_ToolBarControlTypeId: "ToolBar"
  of UIA_ToolTipControlTypeId: "ToolTip"
  of UIA_TreeControlTypeId: "Tree"
  of UIA_TreeItemControlTypeId: "TreeItem"
  of UIA_WindowControlTypeId: "Window"
  else: fmt"ControlType({typeId})"

proc safeCurrentName(element: ptr IUIAutomationElement): string =
  try:
    element.currentName()
  except CatchableError:
    ""

proc safeAutomationId(element: ptr IUIAutomationElement): string =
  try:
    element.currentAutomationId()
  except CatchableError:
    ""

proc safeControlType(element: ptr IUIAutomationElement): int =
  try:
    element.currentControlType()
  except CatchableError:
    -1

proc safeNativeWindowHandle(element: ptr IUIAutomationElement): int =
  try:
    element.nativeWindowHandle()
  except CatchableError:
    0

proc safeRuntimeId(element: ptr IUIAutomationElement): string =
  var arr: ptr SAFEARRAY
  let hr = element.GetRuntimeId(addr arr)
  if FAILED(hr) or arr.isNil:
    return ""

  defer: discard SafeArrayDestroy(arr)

  var lbound, ubound: LONG
  if FAILED(SafeArrayGetLBound(arr, 1, addr lbound)) or FAILED(SafeArrayGetUBound(arr, 1, addr ubound)):
    return ""

  var parts: seq[string] = @[]
  var idx = lbound
  while idx <= ubound:
    var value: LONG
    if SUCCEEDED(SafeArrayGetElement(arr, addr idx, addr value)):
      parts.add($value)
    else:
      parts.add("?")
    inc idx

  "[" & parts.join(",") & "]"

proc parseBoolParam(params: Table[string, string], key: string,
    default = false): bool =
  if key in params:
    let normalized = params[key].toLowerAscii()
    normalized in ["1", "true", "yes", "on"]
  else:
    default

proc selectorForElement(element: ptr IUIAutomationElement, runtimeId = ""): string =
  let automationId = safeAutomationId(element)
  let name = safeCurrentName(element)
  let ctrlType = controlTypeName(safeControlType(element))
  let finalRuntimeId = if runtimeId.len > 0: runtimeId else: safeRuntimeId(element)

  var parts: seq[string] = @[]
  if finalRuntimeId.len > 0:
    parts.add("runtimeId=" & finalRuntimeId)
  if automationId.len > 0:
    parts.add("automationId=\"" & automationId & "\"")
  if name.len > 0:
    parts.add("name=\"" & name & "\"")
  parts.add("controlType=" & ctrlType)
  parts.join(", ")

proc copyToClipboard(text: string, logger: Logger) =
  let sizeBytes = (text.len + 1) * int(sizeof(WCHAR))
  let hMem = GlobalAlloc(GMEM_MOVEABLE, SIZE_T(sizeBytes))
  if hMem == 0:
    if logger != nil:
      logger.error("Failed to allocate clipboard buffer")
    return

  let buffer = GlobalLock(hMem)
  if buffer.isNil:
    discard GlobalFree(hMem)
    if logger != nil:
      logger.error("Failed to lock clipboard buffer")
    return

  # Copy UTF-16 text into the allocated buffer
  let wide = newWideCString(text)
  copyMem(buffer, unsafeAddr wide[0], sizeBytes)
  discard GlobalUnlock(hMem)

  if OpenClipboard(0) == 0:
    discard GlobalFree(hMem)
    if logger != nil:
      logger.error("Failed to open clipboard")
    return

  discard EmptyClipboard()
  if SetClipboardData(CF_UNICODETEXT, hMem) == 0:
    discard GlobalFree(hMem)
    if logger != nil:
      logger.error("Failed to set clipboard data")
  discard CloseClipboard()

proc formatElementInfo(element: ptr IUIAutomationElement): seq[(string, string)] =
  let name = safeCurrentName(element)
  let automationId = safeAutomationId(element)
  let ctrlType = safeControlType(element)
  let hwnd = safeNativeWindowHandle(element)
  let runtimeId = safeRuntimeId(element)

  var fields: seq[(string, string)] = @[("controlType", controlTypeName(ctrlType))]
  if name.len > 0:
    fields.add(("name", name))
  if automationId.len > 0:
    fields.add(("automationId", automationId))
  if runtimeId.len > 0:
    fields.add(("runtimeId", runtimeId))
  if hwnd != 0:
    fields.add(("hwnd", fmt"0x{cast[uint](hwnd):X}"))
  let selector = selectorForElement(element, runtimeId)
  if selector.len > 0:
    fields.add(("selector", selector))
  fields

proc matchesElement(element: ptr IUIAutomationElement, automationId: string,
    controlTypes: openArray[int]): bool =
  let ctrlType = safeControlType(element)
  let idMatch = automationId.len == 0 or safeAutomationId(element) == automationId
  var typeMatch = controlTypes.len == 0
  if not typeMatch:
    for ct in controlTypes:
      if ct == ctrlType:
        typeMatch = true
        break
  idMatch and typeMatch

proc findElement(uia: Uia, element: ptr IUIAutomationElement,
    walker: ptr IUIAutomationTreeWalker, automationId: string,
    controlTypes: openArray[int], depth, maxDepth: int, logger: Logger): ptr IUIAutomationElement =
  if element.isNil or depth > maxDepth:
    return nil

  if matchesElement(element, automationId, controlTypes):
    discard element.AddRef()
    return element

  var child: ptr IUIAutomationElement
  let hrFirst = walker.GetFirstChildElement(element, addr child)
  if FAILED(hrFirst):
    if logger != nil:
      logger.warn(
        "Failed to enumerate UIA children",
        [("depth", $depth), ("hresult", fmt"0x{hrFirst:X}")]
      )
    return nil
  if hrFirst == S_FALSE or child.isNil:
    return nil

  var current = child
  while current != nil:
    let found = findElement(uia, current, walker, automationId, controlTypes, depth + 1, maxDepth, logger)
    if found != nil:
      discard current.Release()
      return found

    var next: ptr IUIAutomationElement
    let hrNext = walker.GetNextSiblingElement(current, addr next)
    discard current.Release()
    if FAILED(hrNext):
      if logger != nil:
        logger.warn(
          "Failed to enumerate UIA siblings",
          [("depth", $depth), ("hresult", fmt"0x{hrNext:X}")]
        )
      break
    if hrNext == S_FALSE:
      break
    current = next

  nil

proc logElementTree(element: ptr IUIAutomationElement,
    walker: ptr IUIAutomationTreeWalker, depth, maxDepth: int,
    logger: Logger) =
  if element.isNil or depth > maxDepth:
    return

  if logger != nil:
    var fields = formatElementInfo(element)
    fields.insert(("depth", $depth), 0)
    logger.info("UIA tree node", fields)

  if depth == maxDepth:
    return

  var child: ptr IUIAutomationElement
  let hrFirst = walker.GetFirstChildElement(element, addr child)
  if FAILED(hrFirst):
    if logger != nil:
      logger.warn(
        "Failed to enumerate UIA children",
        [("depth", $depth), ("hresult", fmt"0x{hrFirst:X}")]
      )
    return
  if hrFirst == S_FALSE or child.isNil:
    return

  var current = child
  while current != nil:
    logElementTree(current, walker, depth + 1, maxDepth, logger)

    var next: ptr IUIAutomationElement
    let hrNext = walker.GetNextSiblingElement(current, addr next)
    discard current.Release()
    if FAILED(hrNext):
      if logger != nil:
        logger.warn(
          "Failed to enumerate UIA siblings",
          [("depth", $depth), ("hresult", fmt"0x{hrNext:X}")]
        )
      break
    if hrNext == S_FALSE:
      break
    current = next

proc resolveRootElement(uia: Uia, params: Table[string, string],
    ctx: RuntimeContext, logger: Logger): ptr IUIAutomationElement =
  let source = params.getOrDefault("source", "target").toLowerAscii()

  if source == "mouse":
    let pt = winMouse.getMousePos()
    try:
      return uia.fromPoint(int32(pt.x), int32(pt.y))
    except CatchableError as exc:
      if logger != nil:
        logger.error("UIA fromPoint failed", [("error", exc.msg)])
      return nil

  var hwnd: int = 0
  let targetName = params.getOrDefault("target", "")
  if targetName.len > 0 and targetName in ctx.windowTargets:
    let target = ctx.windowTargets[targetName]
    if target.storedHwnd.isSome:
      hwnd = target.storedHwnd.get()

  if hwnd == 0:
    try:
      hwnd = ctx.backend.getActiveWindow()
    except CatchableError as exc:
      if logger != nil:
        logger.warn("Failed to read active window", [("error", exc.msg)])

  if hwnd == 0:
    return nil

  try:
    uia.fromWindowHandle(HWND(hwnd))
  except CatchableError as exc:
    if logger != nil:
      logger.error(
        "UIA lookup from HWND failed",
        [("hwnd", $hwnd), ("error", exc.msg)]
      )
    nil

method install*(plugin: UiaPlugin, registry: var ActionRegistry,
    ctx: var RuntimeContext) =
  # Initialize the UIA session for this plugin.
  plugin.uia = initUia()

  # Capture the runtime context by value to avoid escaping a var parameter.
  let ctxValue = ctx

  # Register any UIA-based actions here using `registry.registerAction(...)`
  # (click-button, dump-under-mouse, etc.)
  registry.registerAction("invoke", proc(params: Table[string, string],
      ctx: var RuntimeContext): TaskAction =
    let automationId = params.getOrDefault("automation_id", "")
    let controlTypeParam = params.getOrDefault("control_type", "")
    var controlTypes: seq[int] = @[]
    let parsedControlType = controlTypeFromString(controlTypeParam)
    if parsedControlType != -1:
      controlTypes.add(parsedControlType)

    let logger = ctxValue.logger
    let uia = plugin.uia
    return proc() =
      if uia.isNil:
        if logger != nil:
          logger.error("UIA plugin not initialized")
        return

      var root = resolveRootElement(uia, params, ctxValue, logger)
      if root.isNil:
        if logger != nil:
          logger.warn("No UIA element available for invoke")
        return
      defer: discard root.Release()

      var target = root
      if automationId.len > 0 or controlTypes.len > 0:
        var walker: ptr IUIAutomationTreeWalker
        let hrWalker = uia.automation.get_RawViewWalker(addr walker)
        if FAILED(hrWalker) or walker.isNil:
          if logger != nil:
            logger.error("Failed to create UIA walker", [("hresult", fmt"0x{hrWalker:X}")])
        else:
          defer: discard walker.Release()
          let found = findElement(uia, root, walker, automationId, controlTypes, 0, 12, logger)
          if found != nil:
            target = found
            defer: discard found.Release()

      try:
        uia.invoke(target)
        if logger != nil:
          logger.info("Invoked UIA element", formatElementInfo(target))
      except CatchableError as exc:
        if logger != nil:
          logger.error("UIA invoke failed", [("error", exc.msg)])
  )

  registry.registerAction("uia_dump_element", proc(params: Table[string, string],
      ctx: var RuntimeContext): TaskAction =
    let automationId = params.getOrDefault("automation_id", "")
    let controlTypeParam = params.getOrDefault("control_type", "")
    var controlTypes: seq[int] = @[]
    let parsedControlType = controlTypeFromString(controlTypeParam)
    if parsedControlType != -1:
      controlTypes.add(parsedControlType)

    let logger = ctxValue.logger
    let uia = plugin.uia
    return proc() =
      if uia.isNil:
        if logger != nil:
          logger.error("UIA plugin not initialized")
        return

      var root = resolveRootElement(uia, params, ctxValue, logger)
      if root.isNil:
        if logger != nil:
          logger.warn("No UIA element available to dump")
        return
      defer: discard root.Release()

      var target = root
      if automationId.len > 0 or controlTypes.len > 0:
        var walker: ptr IUIAutomationTreeWalker
        let hrWalker = uia.automation.get_RawViewWalker(addr walker)
        if FAILED(hrWalker) or walker.isNil:
          if logger != nil:
            logger.error("Failed to create UIA walker", [("hresult", fmt"0x{hrWalker:X}")])
        else:
          defer: discard walker.Release()
          let found = findElement(uia, root, walker, automationId, controlTypes, 0, 8, logger)
          if found != nil:
            target = found
            defer: discard found.Release()

      if logger != nil:
        logger.info("UIA element info", formatElementInfo(target))
  )

  registry.registerAction("uia_dump_tree", proc(params: Table[string, string],
      ctx: var RuntimeContext): TaskAction =
    let maxDepth =
      try:
        parseInt(params.getOrDefault("max_depth", "3"))
      except ValueError:
        3

    let logger = ctxValue.logger
    let uia = plugin.uia
    return proc() =
      if uia.isNil:
        if logger != nil:
          logger.error("UIA plugin not initialized")
        return

      var root = resolveRootElement(uia, params, ctxValue, logger)
      if root.isNil:
        if logger != nil:
          logger.warn("No UIA element available to dump tree")
        return
      defer: discard root.Release()

      var walker: ptr IUIAutomationTreeWalker
      let hrWalker = uia.automation.get_RawViewWalker(addr walker)
      if FAILED(hrWalker) or walker.isNil:
        if logger != nil:
          logger.error("Failed to create UIA walker", [("hresult", fmt"0x{hrWalker:X}")])
        return
      defer: discard walker.Release()

      logElementTree(root, walker, 0, maxDepth, logger)
  )

  registry.registerAction("uia_capture_element", proc(params: Table[string, string],
      ctx: var RuntimeContext): TaskAction =
    let logger = ctxValue.logger
    let uia = plugin.uia
    let copySelector = parseBoolParam(params, "copy_selector", false)
    return proc() =
      if uia.isNil:
        if logger != nil:
          logger.error("UIA plugin not initialized")
        return

      var element = resolveRootElement(uia, params, ctxValue, logger)
      if element.isNil:
        if logger != nil:
          logger.warn("No UIA element available to capture")
        return
      defer: discard element.Release()

      let selector = selectorForElement(element)
      if logger != nil:
        var fields = formatElementInfo(element)
        fields.add(("selector", selector))
        logger.info("Captured UIA element", fields)

      if copySelector:
        copyToClipboard(selector, logger)
  )

method shutdown*(plugin: UiaPlugin, ctx: RuntimeContext) =
  discard ctx
  if not plugin.uia.isNil:
    plugin.uia.shutdown()

