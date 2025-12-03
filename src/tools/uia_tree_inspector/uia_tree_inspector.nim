when defined(windows):
  {.define: wDisableResizer.}

  import std/[options, strformat, strutils, times, threadpool]
  import winim/lean
  import winim/com
  import winim/inc/winuser

  import wNim

  import ../../core/scheduler
  import ../../core/logging
  import ../../features/key_parser
  import ../../features/uia/uia
  import ../../platform/windows/hotkeys

  const
    defaultHotkey = "Ctrl+Shift+I"
    defaultMaxDepth = 4

  type
    CaptureSource = enum
      csMouse, csActiveWindow

    TreeNode = ref object
      controlType: string
      name: string
      automationId: string
      hwnd: int
      children: seq[TreeNode]

    TreeBuildResult = object
      node: Option[TreeNode]
      error: string

  proc controlTypeName(typeId: int): string =
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
    of UIA_RadioButtonControlTypeId: "RadioButton"
    of UIA_ScrollBarControlTypeId: "ScrollBar"
    of UIA_SliderControlTypeId: "Slider"
    of UIA_SplitButtonControlTypeId: "SplitButton"
    of UIA_StatusBarControlTypeId: "StatusBar"
    of UIA_TabControlTypeId: "TabControl"
    of UIA_TabItemControlTypeId: "TabItem"
    of UIA_TableControlTypeId: "Table"
    of UIA_TextControlTypeId: "Text"
    of UIA_TitleBarControlTypeId: "TitleBar"
    of UIA_ToolBarControlTypeId: "ToolBar"
    of UIA_ToolTipControlTypeId: "ToolTip"
    of UIA_TreeControlTypeId: "Tree"
    of UIA_TreeItemControlTypeId: "TreeItem"
    of UIA_WindowControlTypeId: "Window"
    else: fmt"ControlType({typeId})"

  proc variantToString(val: VARIANT): string =
    case val.vt
    of VT_BSTR:
      if val.bstrVal == nil:
        ""
      else:
        $cast[WideCString](val.bstrVal)
    of VT_I4:
      $val.lVal
    of VT_UI4:
      $val.ulVal
    of VT_I2:
      $val.iVal
    of VT_BOOL:
      if val.boolVal == VARIANT_TRUE:
        "true"
      else:
        "false"
    else:
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

  proc safeVariantString(uia: Uia, element: ptr IUIAutomationElement,
      propertyId: PROPERTYID): string =
    try:
      var v = uia.getCurrentPropertyValue(element, propertyId)
      defer: discard VariantClear(addr v)
      variantToString(v)
    except CatchableError:
      ""

  proc buildTreeModel(uia: Uia, element: ptr IUIAutomationElement,
      walker: ptr IUIAutomationTreeWalker, depth, maxDepth: int): TreeNode =
    if element.isNil or depth > maxDepth:
      return nil

    let node = TreeNode(
      controlType: controlTypeName(safeControlType(element)),
      name: safeVariantString(uia, element, NamePropertyId),
      automationId: safeVariantString(uia, element, AutomationIdPropertyId),
      hwnd: safeNativeWindowHandle(element),
      children: @[]
    )

    if depth == maxDepth:
      return node

    var child: ptr IUIAutomationElement
    let hrFirst = walker.GetFirstChildElement(element, addr child)
    if FAILED(hrFirst) or hrFirst == S_FALSE or child.isNil:
      return node

    var current = child
    while current != nil:
      let childNode = buildTreeModel(uia, current, walker, depth + 1, maxDepth)
      if childNode != nil:
        node.children.add(childNode)

      var next: ptr IUIAutomationElement
      let hrNext = walker.GetNextSiblingElement(current, addr next)
      discard current.Release()
      if FAILED(hrNext) or hrNext == S_FALSE:
        break
      current = next

    node

  proc buildTreeAsync(source: CaptureSource, maxDepth: int): TreeBuildResult =
    result = TreeBuildResult(node: none(TreeNode), error: "")

    try:
      let uia = initUia(COINIT_MULTITHREADED)
      defer: uia.shutdown()

      var root: ptr IUIAutomationElement
      case source
      of csMouse:
        var pt: POINT
        if GetCursorPos(addr pt) == 0:
          result.error = "Unable to read cursor position."
          return
        root = uia.fromPoint(pt.x, pt.y)
      of csActiveWindow:
        let hwnd = GetForegroundWindow()
        if hwnd == 0:
          result.error = "No active window available."
          return
        root = uia.fromWindowHandle(hwnd)

      defer:
        if root != nil:
          discard root.Release()

      var walker: ptr IUIAutomationTreeWalker
      let hrWalker = uia.automation.get_RawViewWalker(addr walker)
      if FAILED(hrWalker) or walker.isNil:
        result.error = fmt"Failed to create UIA walker (0x{hrWalker:X})"
        return
      defer: discard walker.Release()

      let built = buildTreeModel(uia, root, walker, 0, maxDepth)
      if built.isNil:
        result.error = "No UIA elements were discovered."
      else:
        result.node = some(built)
    except CatchableError as exc:
      result.error = exc.msg

  proc formatNodeLabel(node: TreeNode): string =
    let hwndText =
      if node.hwnd == 0:
        ""
      else:
        fmt" (0x{cast[uint](node.hwnd):X})"
    let nameText =
      if node.name.len > 0:
        " - " & node.name
      else:
        ""
    let automationText =
      if node.automationId.len > 0:
        fmt" [{node.automationId}]"
      else:
        ""
    fmt"{node.controlType}{nameText}{automationText}{hwndText}"

  proc populateTree(tree: wTreeCtrl, node: TreeNode) =
    tree.deleteAllItems()
    if node.isNil:
      return

    proc addChildren(parent: wTreeItem, children: seq[TreeNode]) =
      for child in children:
        let childItem = tree.appendItem(parent, formatNodeLabel(child))
        if child.children.len > 0:
          addChildren(childItem, child.children)

    let rootItem = tree.addRoot(formatNodeLabel(node))
    addChildren(rootItem, node.children)
    rootItem.expand()

  proc layout(panel: wPanel, heading: wStaticText, sourceLabel: wStaticText,
      mouseOption, windowOption: wRadioButton, hotkeyLabel: wStaticText,
      hotkeyInput: wTextCtrl, registerBtn: wButton, depthLabel: wStaticText,
      depthInput: wTextCtrl, inspectBtn: wButton, tree: wTreeCtrl) =
    let padding = 12
    let spacing = 8
    let (w, h) = panel.getClientSize()

    var y = padding

    heading.move(padding, y)
    let headingSize = heading.getBestSize()
    heading.setSize(w - padding * 2, headingSize.height)
    y += headingSize.height + spacing

    sourceLabel.move(padding, y)
    let sourceLabelSize = sourceLabel.getBestSize()
    sourceLabel.setSize(sourceLabelSize.width, sourceLabelSize.height)
    var radioX = padding + sourceLabelSize.width + spacing
    mouseOption.move(radioX, y)
    let mouseSize = mouseOption.getBestSize()
    mouseOption.setSize(mouseSize.width, mouseSize.height)
    radioX += mouseSize.width + spacing
    windowOption.move(radioX, y)
    let windowSize = windowOption.getBestSize()
    windowOption.setSize(windowSize.width, windowSize.height)
    y += max(sourceLabelSize.height, max(mouseSize.height, windowSize.height)) + spacing

    hotkeyLabel.move(padding, y)
    let hotkeyLabelSize = hotkeyLabel.getBestSize()
    hotkeyLabel.setSize(hotkeyLabelSize.width, hotkeyLabelSize.height)
    var hotkeyX = padding + hotkeyLabelSize.width + spacing
    let hotkeyHeight = hotkeyInput.getBestSize().height
    hotkeyInput.move(hotkeyX, y)
    hotkeyInput.setSize(180, hotkeyHeight)
    hotkeyX += 180 + spacing
    registerBtn.move(hotkeyX, y)
    let regSize = registerBtn.getBestSize()
    registerBtn.setSize(regSize.width, regSize.height)
    y += max(hotkeyHeight, regSize.height) + spacing

    depthLabel.move(padding, y)
    let depthLabelSize = depthLabel.getBestSize()
    depthLabel.setSize(depthLabelSize.width, depthLabelSize.height)
    let depthX = padding + depthLabelSize.width + spacing
    let depthHeight = depthInput.getBestSize().height
    depthInput.move(depthX, y)
    depthInput.setSize(60, depthHeight)
    let inspectX = depthX + 60 + spacing
    inspectBtn.move(inspectX, y)
    let inspectSize = inspectBtn.getBestSize()
    inspectBtn.setSize(inspectSize.width, inspectSize.height)
    y += max(depthHeight, inspectSize.height) + spacing

    tree.move(padding, y)
    let treeHeight = max(padding, h - y - padding)
    tree.setSize(w - padding * 2, treeHeight)

  proc main() =
    let logger = newLogger(llInfo)
    let scheduler = newScheduler(logger)
    var pendingBuild: FlowVar[TreeBuildResult]

    let app = App(wSystemDpiAware)
    let frame = Frame(title="UIA Tree Inspector", size=(720, 520))
    let status = StatusBar(frame)
    let panel = Panel(frame)

    let heading = StaticText(panel,
      label="Capture UI Automation trees without freezing the UI. Use the hotkey or click Refresh to rebuild.")

    let sourceLabel = StaticText(panel, label="Capture source:")
    let mouseOption = RadioButton(panel, label="Mouse cursor", style=wRbGroup)
    mouseOption.setValue(true)
    let windowOption = RadioButton(panel, label="Active window")

    let hotkeyLabel = StaticText(panel, label="Global hotkey:")
    let hotkeyInput = TextCtrl(panel, value=defaultHotkey)
    let registerBtn = Button(panel, label="Register")

    let depthLabel = StaticText(panel, label="Max depth:")
    let depthInput = TextCtrl(panel, value=defaultMaxDepth.intToStr())
    let inspectBtn = Button(panel, label="Refresh now")

    let tree = TreeCtrl(panel, style=wTrHasButtons or wTrLinesAtRoot or wTrFullRowHighlight)
    tree.setMinSize((100, 200))

    layout(panel, heading, sourceLabel, mouseOption, windowOption, hotkeyLabel,
      hotkeyInput, registerBtn, depthLabel, depthInput, inspectBtn, tree)

    proc selectedSource(): CaptureSource =
      if mouseOption.getValue():
        csMouse
      else:
        csActiveWindow

    proc maxDepth(): int =
      try:
        let parsed = parseInt(depthInput.getValue().strip())
        if parsed >= 0:
          parsed
        else:
          defaultMaxDepth
      except ValueError:
        defaultMaxDepth

    proc scheduleTraversal() =
      status.setStatusText("Building UIA tree...")
      pendingBuild = spawn buildTreeAsync(selectedSource(), maxDepth())

    proc handleResult(res: TreeBuildResult) =
      if res.error.len > 0:
        status.setStatusText(res.error)
        tree.deleteAllItems()
        return

      if res.node.isSome:
        populateTree(tree, res.node.get())
        status.setStatusText("UIA tree captured.")
      else:
        status.setStatusText("No UIA data available.")
        tree.deleteAllItems()

    proc ensureHotkeyRegistered() =
      let parsed = parseHotkeyString(hotkeyInput.getValue())
      if parsed.key == 0:
        status.setStatusText("Enter a valid hotkey before registering.")
        return
      try:
        unregisterAllHotkeys()
        discard registerHotkey(parsed.modifiers, parsed.key, proc() =
          discard scheduler.scheduleOnce(initDuration(milliseconds = 0), scheduleTraversal)
        )
        status.setStatusText("Global hotkey registered.")
      except IOError as exc:
        status.setStatusText("Hotkey registration failed: " & exc.msg)

    frame.wEvent_Size do (event: wEvent):
      discard event
      layout(panel, heading, sourceLabel, mouseOption, windowOption, hotkeyLabel,
        hotkeyInput, registerBtn, depthLabel, depthInput, inspectBtn, tree)

    registerBtn.wEvent_Button do ():
      ensureHotkeyRegistered()

    inspectBtn.wEvent_Button do ():
      discard scheduler.scheduleOnce(initDuration(milliseconds = 0), scheduleTraversal)

    frame.wEvent_Timer do (event: wEvent):
      discard event
      scheduler.tick()
      pollHotkeyMessages(scheduler)

      if not pendingBuild.isNil and pendingBuild.isReady():
        let res = ^pendingBuild
        pendingBuild = nil
        handleResult(res)

    frame.wEvent_Close do ():
      frame.stopTimer()
      unregisterAllHotkeys()
      frame.delete()

    frame.center()
    frame.show()
    ensureHotkeyRegistered()
    frame.startTimer(16)
    app.mainLoop()

  when isMainModule:
    main()
else:
  when isMainModule:
    static:
      {.fatal: "uia_tree_inspector is only available on Windows.".}
