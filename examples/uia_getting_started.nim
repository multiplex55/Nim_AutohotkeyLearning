import std/[os, times]
import uia

when not defined(windows):
  echo "UI Automation examples only run on Windows."
else:
  let automation = initUia()
  defer: automation.shutdown()

  let root = automation.rootElement()
  echo "Root element name: ", root.currentName()

  echo "Trying to find a button named 'OK'..."
  let okButtonCond = automation.nameAndControlType("OK", UIA_ButtonControlTypeId)
  let okButton = automation.waitElement(tsDescendants, okButtonCond, 2.seconds)
  if okButton != nil:
    okButton.invoke()
  else:
    echo "No 'OK' button visible in current UI tree."

  echo "Trying to find a text box with AutomationId 'Username' and type into it..."
  let username = automation.waitElement(tsDescendants, automation.automationIdCondition("Username"), 2.seconds)
  if username != nil:
    username.setValue("demo-user")
  else:
    echo "No Username field available."
