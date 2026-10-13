' AL-IDRISI DD2027 - one "last change" date for the whole workbook.
' Paste into the ThisWorkbook module (Alt+F11), then save as .xlsm.
' Writes today's date into the named cell LastChange when any value changes
' on any tab, and again on save (catches fill-only status changes).
' Writes only when the date differs, so Excel's undo list is cleared at most once a day.

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    StampLastChange
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    If Not ThisWorkbook.Saved Then StampLastChange
End Sub

Private Sub StampLastChange()
    Dim c As Range
    On Error GoTo Done
    Set c = ThisWorkbook.Names("LastChange").RefersToRange
    If c.Value = Date Then Exit Sub
    Application.EnableEvents = False
    c.Value = Date
Done:
    Application.EnableEvents = True
End Sub
