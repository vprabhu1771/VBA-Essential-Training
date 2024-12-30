If the `Workbook_Open` event in your VBA code is not working when you open the Excel workbook, it could be due to several reasons. Here's how you can troubleshoot and fix the issue:

---

### **1. Ensure Macros Are Enabled**
- If macros are disabled, the `Workbook_Open` event will not run. 
- **Steps to enable macros**:
  1. Go to **File > Options > Trust Center > Trust Center Settings**.
  2. Navigate to **Macro Settings** and select **Enable all macros**.
  3. Restart the workbook.

---

### **2. Save the Workbook as a Macro-Enabled File**
- Ensure your workbook is saved as a **Macro-Enabled Workbook** (`.xlsm`).
- If it's saved as a regular `.xlsx` file, VBA macros will not run.

---

### **3. Check the Code Is in the Correct Location**
- The `Workbook_Open` event should be placed in the **ThisWorkbook** module.
  1. Open the Visual Basic for Applications editor (`Alt + F11`).
  2. Double-click **ThisWorkbook** in the left-hand project explorer.
  3. Ensure your code is in this module.

---

### **4. Verify the `Workbook_Open` Event**
- The event should look like this:
  ```vba
  Private Sub Workbook_Open()
      MsgBox "file opening"
  End Sub
  ```
- Confirm that there are no syntax errors or additional spaces that could cause issues.

---

### **5. Disable Workbook Protection**
- If the workbook or VBA project is password-protected, it might block macros.
  - Remove any protection and test again.

---

### **6. Test for Application Events**
- If the `Workbook_Open` event is still not working, test whether VBA events are enabled.
  - Open the Immediate Window (`Ctrl + G`) and run:
    ```vba
    Application.EnableEvents = True
    ```
  - Save and reopen the workbook.

---

### **7. Try an Alternative Event**
- If the `Workbook_Open` event fails, use the `Auto_Open` subroutine in a standard module:
  ```vba
  Sub Auto_Open()
      MsgBox "file opening"
  End Sub
  ```

---

### **8. Check Excel Add-ins or Third-Party Tools**
- Some add-ins or third-party tools may interfere with macros. Disable them temporarily:
  1. Go to **File > Options > Add-ins**.
  2. Manage and disable add-ins.

---

### **9. Test in a New Workbook**
- Create a new workbook with a simple `Workbook_Open` event to rule out file corruption.

---

If you've followed these steps and the issue persists, let me know for more advanced troubleshooting!