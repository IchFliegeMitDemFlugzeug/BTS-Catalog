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
