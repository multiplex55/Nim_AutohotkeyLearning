when defined(windows):
  {.define: wDisableResizer.}

  import std/[strformat, times]
  import winim/lean
  import winim/com

  import wNim

  import ../../core/scheduler
  import ../../core/logging
  import ../../features/uia/uia

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

  proc describeElement(uia: Uia, element: ptr IUIAutomationElement): string =
    if element == nil:
      return "No element found under the cursor."

    var nameVar = uia.getCurrentPropertyValue(element, NamePropertyId)
    var classVar = uia.getCurrentPropertyValue(element, ClassNamePropertyId)
    var automationIdVar = uia.getCurrentPropertyValue(element, AutomationIdPropertyId)

    defer:
      discard VariantClear(addr nameVar)
      discard VariantClear(addr classVar)
      discard VariantClear(addr automationIdVar)

    let name = variantToString(nameVar)
    let className = variantToString(classVar)
    let automationId = variantToString(automationIdVar)

    result = fmt"Name: {name}\nClass: {className}\nAutomationId: {automationId}"

  proc inspectAtCursor(uia: Uia, output: wTextCtrl) =
    var pt: POINT
    if GetCursorPos(addr pt) == 0:
      output.setValue("Unable to read cursor position.")
      return

    let element = uia.fromPoint(pt.x, pt.y)
    defer:
      if element != nil:
        discard element.Release()

    output.setValue(describeElement(uia, element))

  proc layout(panel: wPanel, heading: wStaticText,
      inspectBtn: wButton, output: wTextCtrl) =
    let padding = 12
    let (w, h) = panel.getClientSize()

    heading.move(padding, padding)
    let headingSize = heading.getBestSize()
    heading.setSize(w - padding * 2, headingSize.height)

    let headingBottom = padding + headingSize.height

    let inspectBtnSize = inspectBtn.getBestSize()
    inspectBtn.move(padding, headingBottom + padding)
    inspectBtn.setSize(inspectBtnSize.width, inspectBtnSize.height)

    let outputY = headingBottom + padding * 2 + inspectBtnSize.height
    let outputHeight =
      if h - outputY - padding > padding:
        h - outputY - padding
      else:
        padding
    output.move(padding, outputY)
    output.setSize(w - padding * 2, outputHeight)

  proc main() =
    let logger = newLogger(llInfo)
    let scheduler = newScheduler(logger)
    let automation = initUia()

    defer:
      automation.shutdown()

    let app = App(wSystemDpiAware)
    let frame = Frame(title="UIA Tree Inspector", size=(520, 380))
    discard StatusBar(frame)
    let panel = Panel(frame)

    let heading = StaticText(panel,
      label="Inspect UI Automation elements without blocking the UI.")
    let inspectBtn = Button(panel, label="Inspect Cursor Element")
    let output = TextCtrl(panel, style=wTeMultiLine or wTeReadOnly)
    output.setValue("Click 'Inspect Cursor Element' to capture details under the mouse pointer.")

    layout(panel, heading, inspectBtn, output)

    frame.wEvent_Size do (event: wEvent):
      discard event
      layout(panel, heading, inspectBtn, output)

    inspectBtn.wEvent_Button do ():
      discard scheduler.scheduleOnce(initDuration(milliseconds = 0), proc() =
        inspectAtCursor(automation, output)
      )

    frame.wEvent_Timer do (event: wEvent):
      discard event
      scheduler.tick()

    frame.wEvent_Close do ():
      frame.stopTimer()
      frame.delete()

    frame.center()
    frame.show()
    frame.startTimer(16)
    app.mainLoop()

  when isMainModule:
    main()
else:
  when isMainModule:
    static:
      {.fatal: "uia_tree_inspector is only available on Windows.".}
