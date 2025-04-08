1 - Introduction

You can modify the titlebar to whatever text you’d like. We make usre of the Auto_Open Macro in this example.

2 - Modify Excel Titlebar Using VBA

1. Put the following code in a module

```vb
Private Sub auto_open()

 'This Macro will run every time the workbook opens.


Application.Caption = ("AutomateExcel.com")


End Sub
```

2. Close your workbook and re-open it.