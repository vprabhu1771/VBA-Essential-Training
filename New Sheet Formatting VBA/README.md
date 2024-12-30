To achieve this in VBA, you can use the `Workbook_NewSheet` event. This event is triggered whenever a new sheet is added to the workbook. Here’s how you can modify the code to automatically type text in cell `A1`, make it bold, set the font size to 70, and merge four cells:

### VBA Code
```vba
Private Sub Workbook_NewSheet(ByVal Sh As Object)
    With Sh.Range("A1")
        ' Set the value in cell A1
        .Value = "Your Company Name"
        
        ' Set the font to bold and font size to 70
        .Font.Bold = True
        .Font.Size = 70
        
        ' Merge cells A1 to D1
        .Resize(1, 4).Merge
        
        ' Center align the text horizontally and vertically
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub
```

### Steps to Implement
1. Press `Alt + F11` to open the VBA editor.
2. In the VBA editor, double-click `ThisWorkbook` in the Project Explorer.
3. Paste the above code into the `ThisWorkbook` code window.
4. Close the VBA editor.
5. Save the workbook as a macro-enabled file (`*.xlsm`).

### Test the Code
1. Create a new sheet in the workbook.
2. The macro will automatically run, inserting the company name in `A1` with the specified formatting and merged cells.

Let me know if you need further assistance!


-----

To achieve your goal in VBA, you can use the `Workbook_NewSheet` event to automatically set the desired properties for cell `A1` whenever a new sheet is created. Here’s how you can do it:

### VBA Code:
```vba
Private Sub Workbook_NewSheet(ByVal Sh As Object)
    With Sh
        ' Set the value in cell A1
        .Range("A1").Value = "Some Company"
        
        ' Merge cells A1:D1
        .Range("A1:D1").Merge
        
        ' Set font properties
        With .Range("A1").Font
            .Bold = True
            .Size = 70
        End With
        
        ' Center align the text horizontally and vertically
        .Range("A1").HorizontalAlignment = xlCenter
        .Range("A1").VerticalAlignment = xlCenter
    End With
End Sub
```

### Explanation of the Code:
1. **Event Trigger**: The `Workbook_NewSheet` event is triggered automatically when a new sheet is added to the workbook.
2. **Set Text**: The value `"Some Company"` is added to cell `A1`.
3. **Merge Cells**: Cells `A1:D1` are merged to span across four columns.
4. **Font Properties**: The text in `A1` is made bold, and its size is set to 70.
5. **Alignment**: The text is centered both horizontally and vertically within the merged cells.

### Steps to Implement:
1. Open the VBA editor by pressing `Alt + F11`.
2. In the Project Explorer, locate **ThisWorkbook** under your workbook's name.
3. Double-click **ThisWorkbook** to open the code editor.
4. Paste the above code into the editor.
5. Close the VBA editor and save your workbook as a macro-enabled workbook (`.xlsm`).

Now, whenever a new sheet is created in the workbook, the specified formatting will automatically apply to cell `A1` of the new sheet.