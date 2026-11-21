' Workbook entry points for BTS-Catalog.
Option Explicit

Private Sub Workbook_Open()
    ' Create all required worksheets and controls when the file opens.
    BTS_Init
End Sub

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    ' Refresh dependent choices or component names after user input.
    If Application.EnableEvents Then BTS_Changed Sh, Target
End Sub

Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    ' An action in the order sheet removes only the selected order branch.
    If Sh.Name <> BTS_T(8) Then Exit Sub
    If Target.Column <> 12 Or Target.Row < 8 Then Exit Sub
    Cancel = True
    Sh.Activate
    Target.Select
    BTS_DeleteSelected
End Sub
