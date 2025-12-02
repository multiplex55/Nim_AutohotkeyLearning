when defined(windows):
  import std/[strformat, times]
  import winim/lean
  import winim/inc/oleauto

  import wNim/[wApp, wFrame, wPanel, wStaticText, wButton, wTextCtrl, wStatusBar]

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

  proc inspectAtCursor(uia: Uia, output: TextCtrl) =
    var pt: POINT
    if GetCursorPos(addr pt) == 0:
      output.setValue("Unable to read cursor position.")
      return

    let element = uia.fromPoint(pt.x, pt.y)
    defer:
      if element != nil:
        discard element.Release()

    output.setValue(describeElement(uia, element))

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

    panel.layout:
      heading:
        left == panel.left + 12
        top == panel.top + 12
        right == panel.right - 12

      inspectBtn:
        left == heading.left
        top == heading.bottom + 8

      output:
        left == heading.left
        top == inspectBtn.bottom + 12
        right == panel.right - 12
        bottom == panel.bottom - 12

    inspectBtn.wEvent_Button do ():
      discard scheduler.scheduleOnce(milliseconds(0), proc() =
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
    frame.startTimer(0.016)
    app.mainLoop()

  when isMainModule:
    main()
else:
  when isMainModule:
    static:
      {.fatal: "uia_tree_inspector is only available on Windows.".}
