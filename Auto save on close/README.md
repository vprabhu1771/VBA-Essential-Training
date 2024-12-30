To automatically save an Excel workbook when it is closed, you can use the `Workbook_BeforeClose` event in the `ThisWorkbook` module. Here's how to set it up:

### Steps:
1. Open your Excel workbook.
2. Press `Alt + F11` to open the VBA editor.
3. In the VBA editor, locate your workbook in the **Project Explorer** on the left. Double-click on `ThisWorkbook`.
4. Paste the following code into the code window:

```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Save the workbook before closing
    Me.Save
End Sub
```

5. Close the VBA editor by pressing `Alt + Q`.
6. Save your workbook as a macro-enabled workbook (`.xlsm`).

### Explanation:
- The `Workbook_BeforeClose` event triggers whenever the workbook is about to close.
- The `Me.Save` command saves the current workbook.

### Optional Enhancements:
If you want to save the workbook only when there are unsaved changes, modify the code as follows:

```vba
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    If Me.Saved = False Then
        Me.Save
    End If
End Sub
```

This checks if the workbook has unsaved changes before saving.

Now, whenever the workbook is closed, it will automatically save without prompting the user.