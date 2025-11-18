import sys
from PySide6.QtWidgets import QApplication

from runtime import VBContext, VBJsonRuntime
from interpreter import VBInterpreter
from gui import VBForm


FORM_JSON = {
    "form": {
        "name": "FormCustomer",
        "title": "Customer Editor with Browser",
        "size": [1000, 700],
        "layout": "vertical",
        "dataContext": "customer",
        "controls": [
            {"type": "Label", "text": "Name:"},
            {"type": "TextBox", "name": "txtName", "bind": "name"},

            {"type": "Label", "text": "Age:"},
            {"type": "TextBox", "name": "txtAge", "bind": "age"},

            {"type": "Label", "text": "Customers (ListBox):"},
            {"type": "ListBox", "name": "lstCustomers", "bind": "customerNames"},

            {"type": "Label", "text": "Status (ComboBox):"},
            {"type": "ComboBox", "name": "cmbStatus", "bind": "statuses"},

            {"type": "Label", "text": "Embedded browser:"},
            {
                "type": "WebBrowser",
                "name": "wbMain",
                "url": "https://www.python.org",
                "events": {
                    "loadFinished": "wbMain_Loaded"
                }
            },

            {
                "type": "Button",
                "name": "btnSave",
                "text": "Save",
                "events": {"click": "btnSave_Click"}
            },
            {
                "type": "Button",
                "name": "btnGoGoogle",
                "text": "Go to Google",
                "events": {"click": "btnGoGoogle_Click"}
            },
            {
                "type": "Button",
                "name": "btnAlert",
                "text": "JS Alert",
                "events": {"click": "btnAlert_Click"}
            }
        ]
    }
}


VB_SOURCE = """Function GetGreeting()
    GetGreeting = "Hello, " & txtName.Text
End Function

Sub FormCustomer_Load()
    Dim c
    c = JsonNew("Customer")
    JsonSet c, "name", "New Customer"
    JsonSet c, "age", 18

    AppData = JsonNew("Root")
    JsonSet AppData, "customer", c

    ' ListBox items: ["Alice","Bob","Charlie"]
    Dim names
    names = JsonParse("[""Alice"",""Bob"",""Charlie""]")
    JsonSet AppData, "customerNames", names

    ' ComboBox items: ["Active","Inactive","Pending"]
    Dim statuses
    statuses = JsonParse("[""Active"",""Inactive"",""Pending""]")
    JsonSet AppData, "statuses", statuses
End Sub

Sub wbMain_Loaded()
    MsgBox "Browser finished loading: " & wbMain.Url
End Sub

Sub btnGoGoogle_Click()
    wbMain.Navigate "https://www.google.com"
End Sub

Sub btnAlert_Click()
    ' Test both EvalJs and BrowserEvalJs
    wbMain.EvalJs "alert('Hello from EvalJs method!')"
    BrowserEvalJs wbMain, "alert('Hello from BrowserEvalJs builtin!')"
End Sub

Sub btnSave_Click()
    JsonSet AppData, "customer.name", txtName.Text
    JsonSet AppData, "customer.age", txtAge.Text

    Dim age
    age = JsonGet(AppData, "customer.age")

    Dim i
    i = 1
    While i <= 3
        i = i + 1
    Wend

    Dim d
    d = 0
    Do While d < 2
        d = d + 1
    Loop

    Dim k
    k = 0
    Do
        k = k + 1
    Loop While k < 2

    Dim custName, statusText
    custName = lstCustomers.SelectedText
    statusText = cmbStatus.Text

    If age >= 18 Then
        MsgBox GetGreeting() & " (adult, " & age & ")" & _
               " | Selected customer: " & custName & _
               " | Status: " & statusText
    Else
        MsgBox GetGreeting() & " (minor, " & age & ")" & _
               " | Selected customer: " & custName & _
               " | Status: " & statusText
    End If
End Sub
"""


def main():
    app = QApplication(sys.argv)

    ctx = VBContext()
    ctx.set_var("AppData", VBJsonRuntime.json_new("Root"))

    interpreter = VBInterpreter(ctx)
    interpreter.load_source(VB_SOURCE)

    form_def = FORM_JSON["form"]
    window = VBForm(form_def, interpreter, ctx)
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
