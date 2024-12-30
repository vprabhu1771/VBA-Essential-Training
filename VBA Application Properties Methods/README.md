Here is a list of commonly used **properties** and **methods** of the `Application` object in VBA:

### Properties:
1. **ActiveCell** - Represents the active cell in the active workbook.
2. **ActiveSheet** - Represents the active sheet in the active workbook.
3. **ActiveWorkbook** - Represents the active workbook.
4. **Calculation** - Sets or returns the calculation mode (e.g., xlCalculationAutomatic, xlCalculationManual).
5. **CellDragAndDrop** - Enables or disables drag-and-drop functionality.
6. **ClipboardFormats** - Returns the formats available on the Clipboard.
7. **CutCopyMode** - Determines whether a cut or copy mode is active.
8. **DisplayAlerts** - Enables or disables system alerts.
9. **EnableEvents** - Enables or disables VBA events.
10. **Hwnd** - Returns the window handle of the Excel application.
11. **ScreenUpdating** - Enables or disables screen updates for performance.
12. **StartupPath** - Returns the startup path for Excel.
13. **StatusBar** - Sets or gets the status bar text.
14. **Version** - Returns the version of Excel.
15. **Workbooks** - Collection of all the open workbooks.

### Methods:
1. **Calculate** - Forces a recalculation of all open workbooks.
2. **Evaluate** - Converts a string into an object or value.
3. **FindFile** - Displays the Open dialog box to find a file.
4. **GetOpenFilename** - Displays a dialog box to select a file.
5. **GetSaveAsFilename** - Displays a Save As dialog box to specify a file name.
6. **Goto** - Scrolls to a specific range or cell.
7. **OnTime** - Schedules a procedure to run at a specific time.
8. **Quit** - Closes the application.
9. **Run** - Runs a specified macro.
10. **Wait** - Pauses the code execution until a specified time.
11. **Undo** - Reverses the last user action.

### Example:
```vba
Sub Example()
    ' Set application properties
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    ' Display version
    MsgBox "Excel Version: " & Application.Version
    
    ' Perform actions
    Application.Calculate
    Application.Goto Reference:=Range("A1")
    
    ' Restore properties
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub
```

Would you like details on any specific property or method?