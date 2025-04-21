Sub record_invoice()

    Dim invoice_no As Long
    Dim invoice_date As Date
    Dim terms As String
    Dim due_date As Date
    Dim bill_to As String
    Dim address As String
    Dim phone As String
    Dim email As String
    
    Dim targetSheet As Worksheet
    Dim nextRow As Long
    
    ' Set the form sheet (update the name as needed)
    Set formSheet = ThisWorkbook.Sheets("Invoice")
    
    ' Set the target sheet
    Set targetSheet = ThisWorkbook.Sheets("Customer")
    
    ' Get the next empty row in the Customer sheet
    nextRow = targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).Row + 1
    
    ' Capture form data
    invoice_no = Range("F3").Value
    invoice_date = Range("F4").Value
    terms = Range("F5").Value
    due_date = Range("F6").Value
    bill_to = Range("B3").Value
    address = Range("B4").Value
    phone = Range("B5").Value
    email = Range("B6").Value

    ' Write to "Customer" sheet
    With targetSheet
        .Cells(nextRow, 1).Value = invoice_no
        .Cells(nextRow, 2).Value = invoice_date
        .Cells(nextRow, 3).Value = terms
        .Cells(nextRow, 4).Value = due_date
        .Cells(nextRow, 5).Value = bill_to
        .Cells(nextRow, 6).Value = address
        .Cells(nextRow, 7).Value = phone
        .Cells(nextRow, 8).Value = email
    End With
    
    ' Record the order items
    Call record_order_items(invoice_no)

    MsgBox "Invoice recorded successfully.", vbInformation
    
    ' Call clear sub
    ' clear_invoice

End Sub
Sub record_order_items(invoice_no As Long)

    Dim itemSheet As Worksheet
    Dim sourceSheet As Worksheet
    Set itemSheet = ThisWorkbook.Sheets("OrderItems")
    Set sourceSheet = ThisWorkbook.Sheets("Invoice") ' Your form sheet name

    Dim startRow As Long
    startRow = 9 ' Items start from row 9
    
    Dim currentRow As Long
    currentRow = startRow
    
    Dim code As String
    Dim itemName As String
    Dim qty As Long
    Dim unitPrice As Double
    Dim total As Double
    
    ' Find next empty row in OrderItems sheet, starting from row 3 (since headers are in row 2)
    Dim targetRow As Long
    targetRow = itemSheet.Cells(itemSheet.Rows.Count, "A").End(xlUp).Row + 1
    If targetRow < 2 Then targetRow = 2 ' Ensure starting from row 3 if only header exists

    ' Loop until Item Name (Column B) is blank
    Do While Trim(sourceSheet.Cells(currentRow, 2).Value) <> ""

        code = sourceSheet.Cells(currentRow, 1).Value         ' Column A - Code
        itemName = sourceSheet.Cells(currentRow, 2).Value     ' Column B - Item Name
        qty = Val(sourceSheet.Cells(currentRow, 3).Value)     ' Column C - Qty
        unitPrice = Val(sourceSheet.Cells(currentRow, 4).Value) ' Column D - Unit Price
        total = Val(sourceSheet.Cells(currentRow, 5).Value)   ' Column E - Total Price

        ' Write to OrderItems sheet
        With itemSheet
            .Cells(targetRow, 1).Value = invoice_no   ' A: Invoice No
            .Cells(targetRow, 2).Value = code         ' B: Code
            .Cells(targetRow, 3).Value = itemName     ' C: Item Name
            .Cells(targetRow, 4).Value = qty          ' D: Qty
            .Cells(targetRow, 5).Value = unitPrice    ' E: Unit Price
            .Cells(targetRow, 6).Value = total        ' F: Total Price
        End With

        currentRow = currentRow + 1
        targetRow = targetRow + 1
    Loop

End Sub


Sub clear_invoice()

    ' Clear invoice form fields
    With ActiveSheet
        .Range("B3").Value = ""
        .Range("B4").Value = ""
        .Range("B5").Value = ""
        .Range("B6").Value = ""
        .Range("F3").Value = ""
        .Range("F4").Value = ""
        .Range("F5").Value = ""
        .Range("F6").Value = ""
    End With
    
    ' Clear order items table starting from A10
    Range("A9:E100").ClearContents

    MsgBox "Form cleared.", vbInformation

End Sub
