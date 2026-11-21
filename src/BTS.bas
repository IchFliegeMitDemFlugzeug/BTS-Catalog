' BTS-Catalog VBA main module. One record per item and reversible tree views.
Option Explicit

Private Const MAX_DEPTH As Long = 6
Private Const FORM_FIRST As Long = 15
Private Const FORM_LAST As Long = 54

Private Function S(ByVal id As Long) As Worksheet
    Set S = ThisWorkbook.Worksheets(BTS_T(id))
End Function

Private Function DB() As Worksheet
    Set DB = ThisWorkbook.Worksheets("_BTS_DATA")
End Function

Private Function EndRow(ByVal sh As Worksheet, ByVal col As Long) As Long
    EndRow = sh.Cells(sh.Rows.Count, col).End(xlUp).Row
    If EndRow < 1 Then EndRow = 1
End Function

Private Function Locate(ByVal sh As Worksheet, ByVal col As Long, ByVal code As String, Optional ByVal startAt As Long = 2) As Long
    Dim r As Long
    For r = startAt To EndRow(sh, col)
        If StrComp(Trim$(CStr(sh.Cells(r, col).Value)), code, vbTextCompare) = 0 Then
            Locate = r
            Exit Function
        End If
    Next r
    Locate = 0
End Function

Private Function PartRow(ByVal code As String) As Long
    PartRow = Locate(S(2), 1, code)
End Function

Private Function AssemblyRow(ByVal code As String) As Long
    AssemblyRow = Locate(DB, 1, code)
End Function

Private Function StandardRow(ByVal code As String) As Long
    StandardRow = Locate(S(7), 1, code)
End Function

Private Function ExistsCode(ByVal code As String) As Boolean
    ExistsCode = (PartRow(code) > 0 Or AssemblyRow(code) > 0 Or StandardRow(code) > 0)
End Function

Private Function IsAssembly(ByVal code As String) As Boolean
    IsAssembly = AssemblyRow(code) > 0
End Function

Private Function ItemName(ByVal code As String) As String
    Dim p As Long
    p = PartRow(code)
    If p > 0 Then ItemName = CStr(S(2).Cells(p, 2).Value): Exit Function
    p = AssemblyRow(code)
    If p > 0 Then ItemName = CStr(DB.Cells(p, 2).Value): Exit Function
    p = StandardRow(code)
    If p > 0 Then ItemName = CStr(S(7).Cells(p, 2).Value)
End Function

Private Function MaterialMark(ByVal code As String) As String
    Dim p As Long
    p = PartRow(code)
    If p > 0 Then MaterialMark = CStr(S(2).Cells(p, 3).Value)
End Function

Private Function Kind(ByVal code As String) As Long
    If IsAssembly(code) Then Kind = 0: Exit Function
    If PartRow(code) > 0 Then Kind = 1: Exit Function
    Kind = 2
End Function

Private Function NextFree(ByVal sh As Worksheet, ByVal col As Long) As Long
    NextFree = EndRow(sh, col) + 1
    If NextFree < 2 Then NextFree = 2
End Function

Private Sub LayoutHeader(ByVal sh As Worksheet, ByVal headerCells As Variant)
    Dim i As Long
    For i = LBound(headerCells) To UBound(headerCells)
        sh.Cells(1, i + 1).Value = headerCells(i)
    Next i
    With sh.Range(sh.Cells(1, 1), sh.Cells(1, UBound(headerCells) + 1))
        .Font.Name = "Arial"
        .Font.Size = 10
        .Font.Bold = True
        .Interior.Color = RGB(227, 232, 236)
        .RowHeight = 26
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(182, 190, 196)
    End With
    sh.Cells.Font.Name = "Arial"
    sh.Cells.Font.Size = 10
    sh.Activate
    ActiveWindow.FreezePanes = False
    sh.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Sub MakeButton(ByVal sh As Worksheet, ByVal key As String, ByVal label As String, ByVal macroName As String, ByVal atCell As String, Optional ByVal widthPoints As Long = 120)
    Dim obj As Shape
    On Error Resume Next
    Set obj = sh.Shapes(key)
    On Error GoTo 0
    If obj Is Nothing Then
        Set obj = sh.Shapes.AddShape(msoShapeRoundedRectangle, sh.Range(atCell).Left, sh.Range(atCell).Top, widthPoints, 25)
        obj.Name = key
    End If
    obj.OnAction = macroName
    obj.TextFrame.Characters.Text = label
    obj.Fill.ForeColor.RGB = RGB(227, 232, 236)
    obj.Line.ForeColor.RGB = RGB(145, 156, 166)
    obj.TextFrame.Characters.Font.Name = "Arial"
    obj.TextFrame.Characters.Font.Size = 9
    obj.TextFrame.Characters.Font.Color = RGB(30, 43, 52)
End Sub

Private Sub EnsureName(ByVal label As String, ByVal addressRef As String)
    On Error Resume Next
    ThisWorkbook.Names(label).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=label, RefersTo:="=" & addressRef
End Sub

Private Sub PutList(ByVal target As Range, ByVal nameRef As String)
    On Error Resume Next
    target.Validation.Delete
    On Error GoTo 0
    target.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & nameRef
    target.Validation.IgnoreBlank = True
    target.Validation.InCellDropdown = True
End Sub

Public Sub BTS_Init()
    Dim i As Long, sh As Worksheet, dbSh As Worksheet, nm As String
    On Error GoTo Failed
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    For i = 1 To 9
        nm = BTS_T(i)
        Set sh = Nothing
        On Error Resume Next
        Set sh = ThisWorkbook.Worksheets(nm)
        On Error GoTo Failed
        If sh Is Nothing Then
            Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            sh.Name = nm
        End If
    Next i
    Set dbSh = Nothing
    On Error Resume Next
    Set dbSh = ThisWorkbook.Worksheets("_BTS_DATA")
    On Error GoTo Failed
    If dbSh Is Nothing Then
        Set dbSh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        dbSh.Name = "_BTS_DATA"
    End If
    If CStr(dbSh.Range("A1").Value) = "" Then
        dbSh.Range("A1:O1").Value = Array("Assembly", "Name", "Family", "Link", "SubLink", "Parent", "Child", "Qty", "Role", "Reserved", "OrderCode", "OrderQty", "Reserved", "ExcludeRoot", "ExcludePath")
    End If
    BTS_ReferenceInit
    If CStr(S(2).Range("A1").Value) = "" Then LayoutHeader S(2), Array(BTS_T(10), BTS_T(11), BTS_T(12))
    If CStr(S(7).Range("A1").Value) = "" Then LayoutHeader S(7), Array(BTS_T(10), BTS_T(11))
    S(2).Columns("A").ColumnWidth = 37
    S(2).Columns("B").ColumnWidth = 48
    S(2).Columns("C").ColumnWidth = 21
    S(7).Columns("A").ColumnWidth = 53
    S(7).Columns("B").ColumnWidth = 53
    MakeButton S(2), "btnDelete", BTS_T(25), "BTS_DeleteSelected", "E2", 140
    For i = 3 To 6
        If CStr(S(i).Range("A1").Value) = "" Then LayoutHeader S(i), Array("I", "II", "III", "IV", "V", "VI", BTS_T(11), BTS_T(12), BTS_T(13))
        S(i).Columns("A:F").ColumnWidth = 29
        S(i).Columns("G").ColumnWidth = 42
        S(i).Columns("H").ColumnWidth = 19
        S(i).Columns("I").ColumnWidth = 11
        S(i).Columns("J:K").Hidden = True
        MakeButton S(i), "btnDelete", BTS_T(25), "BTS_DeleteSelected", "M2", 140
    Next i
    BTS_FormInit
    BTS_OrderInit
    dbSh.Visible = xlSheetVeryHidden
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name = "Sheet1" Or sh.Name = ChrW(1051) & ChrW(1080) & ChrW(1089) & ChrW(1090) & "1" Then
            If WorksheetFunction.CountA(sh.Cells) = 0 And ThisWorkbook.Worksheets.Count > 10 Then
                Application.DisplayAlerts = False
                sh.Delete
                Application.DisplayAlerts = True
            End If
        End If
    Next sh
    BTS_RebuildAll
Clean:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub
Failed:
    MsgBox "BTS_Init: " & Err.Description, vbExclamation
    Resume Clean
End Sub

Private Sub BTS_ReferenceInit()
    ' Reference sheet uses a narrow vertical layout, never extra columns N and beyond.
    Dim sh As Worksheet, i As Long
    Set sh = S(9)
    If CStr(sh.Range("A1").Value) = "" Then
        sh.Range("A1").Value = BTS_T(40)
        sh.Range("B1").Value = BTS_T(14)
        For i = 0 To 3
            sh.Cells(i + 2, 1).Value = Choose(i + 1, "SH", "KR", "DU", "DT")
            sh.Cells(i + 2, 2).Value = BTS_T(i + 3)
        Next i
        sh.Range("A8").Value = BTS_T(40)
        sh.Range("B8").Value = BTS_T(12)
        sh.Range("A16").Value = BTS_T(40)
        sh.Range("B16").Value = BTS_T(69)
        sh.Range("C16").Value = BTS_T(12)
        For i = 0 To 3
            sh.Cells(i + 9, 1).Value = Choose(i + 1, "AL", "BR", "LA", "ST")
            sh.Cells(i + 17, 1).Value = Choose(i + 1, "AL", "BR", "LA", "ST")
            sh.Cells(i + 17, 2).Value = BTS_T(i + 61)
            sh.Cells(i + 17, 3).Value = Choose(i + 1, "AL", "BR", "LA", "ST")
        Next i
        sh.Range("A27").Value = BTS_T(40)
        sh.Range("B27").Value = BTS_T(17)
        sh.Range("A37").Value = BTS_T(40)
        sh.Range("B37").Value = BTS_T(18)
        For i = 0 To 4
            sh.Cells(i + 28, 1).Value = Choose(i + 1, "40", "45", "50", "55", "60")
            sh.Cells(i + 38, 1).Value = Choose(i + 1, "00", "30", "45", "60", "90")
        Next i
        sh.Range("A47").Value = BTS_T(10)
        sh.Range("B47").Value = "BTS.XX.MM.DDAA.NNN"
    End If
    sh.Columns("A").ColumnWidth = 19
    sh.Columns("B").ColumnWidth = 37
    sh.Columns("C").ColumnWidth = 24
    EnsureName "BTSGroups", "'" & sh.Name & "'!$A$2:$A$5"
    EnsureName "BTSMats", "'" & sh.Name & "'!$A$9:$A$15"
    EnsureName "BTSMarks", "'" & sh.Name & "'!$B$17:$B$26"
    EnsureName "BTSDiam", "'" & sh.Name & "'!$A$28:$A$36"
    EnsureName "BTSAngles", "'" & sh.Name & "'!$A$38:$A$46"
End Sub

Private Sub BTS_FormInit()
    ' The form layout is independent of the stored parts and assembly records.
    Dim sh As Worksheet, key As Variant, fresh As Boolean, r As Long
    Set sh = S(1)
    fresh = (CStr(DB.Range("AL1").Value) <> "FORM3" Or CStr(sh.Range("B1").Value) <> BTS_T(72))
    If fresh Then
        sh.Columns("A:I").Hidden = False
        sh.Rows("1:65").Hidden = False
        sh.Range("A1:I65").UnMerge
        sh.Range("A1:I65").Clear
        sh.Range("C5:G5").Merge
        sh.Range("C6:G6").Merge
        sh.Range("C7:E7").Merge
        sh.Range("C9:D9").Merge
        sh.Range("C10:D10").Merge
        sh.Range("C11:D11").Merge
        sh.Range("B1").Value = BTS_T(72)
        sh.Range("C1").Value = BTS_T(73)
        sh.Range("G4").NumberFormat = "@"
        sh.Range("G4").Value = "001"
        DB.Range("AL1").Value = "FORM3"
    End If
    sh.Range("B3:G3").Value = Array(BTS_T(14), BTS_T(12), BTS_T(16), BTS_T(17), BTS_T(18), BTS_T(19))
    sh.Range("B5").Value = BTS_T(10)
    sh.Range("B6").Value = BTS_T(11)
    sh.Range("B7").Value = BTS_T(69)
    sh.Range("B9").Value = BTS_T(78)
    sh.Range("C9").Value = BTS_T(10)
    sh.Range("E9").Value = BTS_T(13)
    sh.Range("B10").Value = BTS_T(79)
    sh.Range("B11").Value = BTS_T(70)
    sh.Range("B13").Value = BTS_T(20)
    sh.Range("A14:F14").Value = Array(BTS_T(80), BTS_T(10), BTS_T(11), BTS_T(13), BTS_T(21), BTS_T(81))
    sh.Columns("A").ColumnWidth = 6
    sh.Columns("B").ColumnWidth = 37
    sh.Columns("C").ColumnWidth = 34
    sh.Columns("D").ColumnWidth = 17
    sh.Columns("E").ColumnWidth = 19
    sh.Columns("F:G").ColumnWidth = 13
    With sh.Range("A1:G54")
        .Font.Name = "Arial"
        .Font.Size = 10
        .VerticalAlignment = xlCenter
        .WrapText = True
        .RowHeight = 22
    End With
    With sh.Range("B3:G3,B9:E9,A14:F14")
        .Font.Bold = True
        .Interior.Color = RGB(227, 232, 236)
        .RowHeight = 32
    End With
    With sh.Range("B1:C1,B3:G4,B5:G7,B9:E11,A14:F54").Borders
        .LineStyle = xlContinuous
        .Color = RGB(175, 183, 190)
        .Weight = xlThin
    End With
    sh.Rows("8:8").RowHeight = 30
    sh.Rows("10:11").RowHeight = 31
    sh.Rows("15:54").RowHeight = 30
    sh.Range("E4:F4").NumberFormat = "00"
    sh.Range("G4").NumberFormat = "@"
    sh.Range("C5:G5").NumberFormat = "@"
    sh.Range("C10:D11").NumberFormat = "@"
    sh.Range("B15:B54").NumberFormat = "@"
    sh.Range("D15:D54").NumberFormat = "0"
    sh.Range("A14:A54,D14:D54").HorizontalAlignment = xlCenter
    DB.Range("AG1:AG2").Value = Application.Transpose(Array(BTS_T(73), BTS_T(74)))
    EnsureName "BTSItemKinds", "'_BTS_DATA'!$AG$1:$AG$2"
    PutList sh.Range("C1"), "BTSItemKinds"
    PutList sh.Range("B4"), "BTSGroups"
    PutList sh.Range("C4"), "BTSMats"
    PutList sh.Range("E4"), "BTSDiam"
    PutList sh.Range("F4"), "BTSAngles"
    ' Draft rows are plain editable cells, not dropdown fields.
    sh.Range("A15:F54").Validation.Delete
    For r = 1 To 100
        DB.Cells(r + 1, 37).Value = r
    Next r
    EnsureName "BTSDraftQty", "'_BTS_DATA'!$AK$2:$AK$101"
    PutList sh.Range("E10:E11"), "BTSDraftQty"
    sh.Range("E10:E11").Validation.ShowError = False
    If sh.Range("E10").Value = "" Then sh.Range("E10").Value = 1
    If sh.Range("E11").Value = "" Then sh.Range("E11").Value = 1
    For Each key In Array("addItem", "saveItem", "loadItem", "propose", "draftOwn", "draftStd")
        On Error Resume Next
        sh.Shapes(CStr(key)).Delete
        On Error GoTo 0
    Next key
    MakeButton sh, "addItem", BTS_T(26), "BTS_AddItem", "B8", 145
    MakeButton sh, "draftOwn", BTS_T(82), "BTS_AddDraftOwn", "F10", 118
    MakeButton sh, "draftStd", BTS_T(82), "BTS_AddDraftStandard", "F11", 118
    BTS_ComponentChoices
    BTS_FormMode fresh
    BTS_RefreshPreview
End Sub

Private Function BTS_FormIsAssembly() As Boolean
    BTS_FormIsAssembly = (CStr(S(1).Range("C1").Value) = BTS_T(74))
End Function

Private Sub BTS_MarkChoices()
    ' Material grades belong to a single registered material family.
    Dim sh As Worksheet, ref As Worksheet, data As Worksheet, code As String, r As Long, p As Long
    Set sh = S(1)
    Set ref = S(9)
    Set data = DB
    code = CStr(sh.Range("C4").Value)
    data.Range("AH2:AH1000").ClearContents
    p = 2
    For r = 17 To 26
        If code <> "" And StrComp(Trim$(CStr(ref.Cells(r, 1).Value)), code, vbTextCompare) = 0 Then
            If Trim$(CStr(ref.Cells(r, 2).Value)) <> "" Then
                data.Cells(p, 34).Value = ref.Cells(r, 2).Value
                p = p + 1
            End If
        End If
    Next r
    EnsureName "BTSSelectedMarks", "'_BTS_DATA'!$AH$2:$AH$" & CStr(Application.Max(2, p - 1))
    PutList sh.Range("C7:E7"), "BTSSelectedMarks"
    If Not ValidMark(CStr(sh.Range("C7").Value), code) Then sh.Range("C7:E7").ClearContents
End Sub

Private Sub BTS_FormMode(ByVal resetEntry As Boolean)
    Dim sh As Worksheet, isUnit As Boolean
    Set sh = S(1)
    isUnit = BTS_FormIsAssembly()
    If resetEntry Then
        sh.Range("D4,G4,C5:G5,C6:G6,C7:E7,C10:D11,A15:F54").ClearContents
        sh.Range("E10:E11").Value = 1
        If isUnit Then
            sh.Range("G4").Value2 = "000"
        Else
            sh.Range("G4").Value2 = "001"
        End If
    End If
    sh.Columns("D").Hidden = Not isUnit
    sh.Rows("7").Hidden = isUnit
    sh.Rows("9:54").Hidden = Not isUnit
    sh.Shapes("draftOwn").Visible = IIf(isUnit, msoTrue, msoFalse)
    sh.Shapes("draftStd").Visible = IIf(isUnit, msoTrue, msoFalse)
    If isUnit Then
        sh.Range("C3").Value = BTS_T(75)
        sh.Range("D3").Value = BTS_T(76)
        PutList sh.Range("D4"), "BTSMats"
        sh.Range("C7:E7").ClearContents
        sh.Range("C7:E7").Validation.Delete
    Else
        sh.Range("C3").Value = BTS_T(12)
        sh.Range("D4").ClearContents
        sh.Range("D4").Validation.Delete
        BTS_MarkChoices
    End If
    BTS_SortDraft
End Sub

Private Sub BTS_RefreshPreview()
    Dim sh As Worksheet, code As String
    Set sh = S(1)
    code = CodeFrom(sh)
    sh.Range("C5").Value = code
    If BTS_FormIsAssembly() And CStr(sh.Range("B4").Value) = "SH" Then
        sh.Range("C6").Value = AutomaticName(code)
    End If
End Sub

Private Sub BTS_OrderInit()
    Dim sh As Worksheet
    Set sh = S(8)
    If CStr(sh.Range("A1").Value) = "" Then LayoutHeader sh, Array(BTS_T(14), BTS_T(15), BTS_T(16), BTS_T(17), BTS_T(18), BTS_T(19), BTS_T(10), BTS_T(13))
    sh.Columns("A:F").ColumnWidth = 16
    sh.Columns("G").ColumnWidth = 40
    sh.Columns("H").ColumnWidth = 17
    sh.Range("D2:E2").NumberFormat = "00"
    sh.Range("F2").NumberFormat = "@"
    sh.Range("G2").Formula = "=IFERROR(IF(OR(A2="""",B2="""",D2="""",E2="""",F2="""",AND(MOD(VALUE(F2),10)=0,C2="""")),"""",""BTS.""&A2&"".""&B2&IF(MOD(VALUE(F2),10)=0,C2,"""")&"".""&TEXT(VALUE(D2),""00"")&TEXT(VALUE(E2),""00"")&"".""&TEXT(VALUE(F2),""000"")),"""")"
    If IsEmpty(sh.Range("H2").Value) Then sh.Range("H2").Value = 1
    sh.Range("A7:I7").Value = Array("I", "II", "III", "IV", "V", "VI", BTS_T(11), BTS_T(12), BTS_T(13))
    sh.Range("L7").Value = BTS_T(25)
    sh.Columns("L").ColumnWidth = 18
    sh.Range("A7:I7").Font.Bold = True
    sh.Range("A7:I7").Interior.Color = RGB(227, 232, 236)
    sh.Columns("A:F").ColumnWidth = 29
    sh.Columns("G").ColumnWidth = 44
    sh.Columns("H").ColumnWidth = 20
    sh.Columns("I").ColumnWidth = 14
    sh.Columns("J:K").Hidden = True
    MakeButton sh, "orderAdd", BTS_T(28), "BTS_AddToOrder", "L2", 120
    MakeButton sh, "orderClear", BTS_T(29), "BTS_ClearOrder", "N2", 120
    MakeButton sh, "orderExport", BTS_T(30), "BTS_ExportOrder", "P2", 148
    MakeButton sh, "orderDelete", BTS_T(25), "BTS_DeleteSelected", "L7", 148
    BTS_UpdateOrderLists
End Sub

Private Function CodeFrom(ByVal sh As Worksheet) As String
    ' Order keeps its original A2:F2 form; creation uses B4:G4.
    Dim a As String, b As String, c As String, d As String, e As String, f As String, rr As Long, offset As Long, isUnit As Boolean
    rr = 2
    If sh.Name = BTS_T(1) Then
        rr = 4
        offset = 1
    End If
    a = Trim$(CStr(sh.Cells(rr, 1 + offset).Value2))
    b = Trim$(CStr(sh.Cells(rr, 2 + offset).Value2))
    c = Trim$(CStr(sh.Cells(rr, 3 + offset).Value2))
    d = Trim$(CStr(sh.Cells(rr, 4 + offset).Value2))
    e = Trim$(CStr(sh.Cells(rr, 5 + offset).Value2))
    f = Trim$(CStr(sh.Cells(rr, 6 + offset).Value2))
    If a = "" Or Len(b) <> 2 Or d = "" Or e = "" Or f = "" Then Exit Function
    If Len(f) > 3 Or Len(d) > 2 Or Len(e) > 2 Then Exit Function
    If f Like "*[!0-9]*" Or d Like "*[!0-9]*" Or e Like "*[!0-9]*" Then Exit Function
    isUnit = (CLng(f) Mod 10 = 0)
    If rr = 4 Then
        If CStr(sh.Range("C1").Value) <> BTS_T(73) And CStr(sh.Range("C1").Value) <> BTS_T(74) Then Exit Function
        If isUnit <> BTS_FormIsAssembly() Then Exit Function
    End If
    If isUnit Then
        If Len(c) <> 2 Then Exit Function
        b = b & c
    End If
    CodeFrom = "BTS." & a & "." & b & "." & Format$(CLng(d), "00") & Format$(CLng(e), "00") & "." & Format$(CLng(f), "000")
End Function

Private Function IsAssemblyCode(ByVal code As String) As Boolean
    Dim ar As Variant
    ar = Split(code, ".")
    If UBound(ar) <> 4 Then Exit Function
    If Not IsNumeric(ar(4)) Then Exit Function
    IsAssemblyCode = CLng(ar(4)) Mod 10 = 0
End Function

Private Function AutomaticName(ByVal code As String) As String
    Dim ar As Variant, ang As Long, diam As String
    ar = Split(code, ".")
    If UBound(ar) <> 4 Then Exit Function
    If ar(1) <> "SH" Then Exit Function
    diam = CStr(Val(Left$(ar(3), 2)) / 10)
    diam = Replace(diam, ".", ",")
    ang = CLng(Val(Right$(ar(3), 2)))
    If ang = 0 Then
        AutomaticName = BTS_T(31) & diam & BTS_T(34)
    Else
        AutomaticName = BTS_T(32) & CStr(ang) & BTS_T(33) & diam & BTS_T(34)
    End If
End Function

Private Function ValidMark(ByVal markName As String, ByVal matCode As String) As Boolean
    Dim sh As Worksheet, rowN As Long
    Set sh = S(9)
    If Trim$(markName) = "" Or Trim$(matCode) = "" Then Exit Function
    For rowN = 17 To 26
        If StrComp(Trim$(CStr(sh.Cells(rowN, 2).Value)), Trim$(markName), vbTextCompare) = 0 And StrComp(Trim$(CStr(sh.Cells(rowN, 1).Value)), Trim$(matCode), vbTextCompare) = 0 Then
            ValidMark = True
            Exit Function
        End If
    Next rowN
End Function

Private Function HasPath(ByVal fromCode As String, ByVal targetCode As String, ByVal depth As Long) As Boolean
    Dim data As Worksheet, i As Long, component As String
    If fromCode = targetCode Then HasPath = True: Exit Function
    If depth > 12 Then HasPath = True: Exit Function
    Set data = DB
    For i = 2 To EndRow(data, 6)
        If CStr(data.Cells(i, 6).Value) = fromCode Then
            component = CStr(data.Cells(i, 7).Value)
            If HasPath(component, targetCode, depth + 1) Then HasPath = True: Exit Function
        End If
    Next i
End Function

Private Function ValidDraftQty(ByVal q As Variant) As Boolean
    ' Guard against blanks, text, fractions, and overflow of the stored integer.
    Dim amount As Double
    If IsError(q) Then Exit Function
    If Not IsNumeric(q) Then Exit Function
    amount = CDbl(q)
    If amount < 1 Or amount > 2147483647# Then Exit Function
    ValidDraftQty = (Fix(amount) = amount)
End Function

Private Function ValidateForm(ByVal code As String, ByVal overwrite As Boolean) As Boolean
    Dim sh As Worksheet, r As Long, cmp As String, q As Variant, nBase As Long, nHead As Long, ar As Variant, baseMat As String, headMat As String, kindAssembly As Boolean, seen As Object, role As String
    Set sh = S(1)
    If code = "" Then MsgBox BTS_T(77), vbExclamation: Exit Function
    If overwrite Then
        If Not ExistsCode(code) Then MsgBox BTS_T(41), vbExclamation: Exit Function
    Else
        If ExistsCode(code) Then MsgBox BTS_T(35), vbExclamation: Exit Function
    End If
    kindAssembly = IsAssemblyCode(code)
    ar = Split(code, ".")
    If kindAssembly Then
        Set seen = CreateObject("Scripting.Dictionary")
        seen.CompareMode = vbTextCompare
        If Len(ar(2)) <> 4 Then MsgBox BTS_T(42), vbExclamation: Exit Function
        baseMat = Left$(ar(2), 2)
        headMat = Right$(ar(2), 2)
        For r = FORM_FIRST To FORM_LAST
            cmp = Trim$(CStr(sh.Cells(r, 2).Value))
            If cmp <> "" Then
                If seen.Exists(cmp) Then MsgBox BTS_T(43), vbExclamation: Exit Function
                seen.Add cmp, True
                q = sh.Cells(r, 4).Value
                If Not ExistsCode(cmp) Or Not ValidDraftQty(q) Then MsgBox BTS_T(43), vbExclamation: Exit Function
                If cmp = code Or HasPath(cmp, code, 0) Then MsgBox BTS_T(44), vbExclamation: Exit Function
                role = Trim$(CStr(sh.Cells(r, 5).Value))
                If StrComp(role, BTS_T(22), vbTextCompare) = 0 Then
                    nBase = nBase + 1
                    If ar(1) = "SH" And ComponentMaterial(cmp) <> baseMat Then MsgBox BTS_T(45), vbExclamation: Exit Function
                ElseIf StrComp(role, BTS_T(23), vbTextCompare) = 0 Then
                    nHead = nHead + 1
                    If ar(1) = "SH" And ComponentMaterial(cmp) <> headMat Then MsgBox BTS_T(45), vbExclamation: Exit Function
                End If
            End If
        Next r
        ' One designation with quantity > 1 is still one composition position.
        If seen.Count < 2 Then MsgBox BTS_T(83), vbExclamation: Exit Function
        If ar(1) = "SH" Then
            If nBase <> 1 Or nHead <> 1 Then MsgBox BTS_T(46), vbExclamation: Exit Function
        End If
    Else
        If Trim$(CStr(sh.Range("C6").Value)) = "" Then MsgBox BTS_T(47), vbExclamation: Exit Function
        If Not ValidMark(CStr(sh.Range("C7").Value), CStr(ar(2))) Then MsgBox BTS_T(48), vbExclamation: Exit Function
    End If
    ValidateForm = True
End Function

Private Function ComponentMaterial(ByVal code As String) As String
    Dim fields As Variant
    fields = Split(code, ".")
    If UBound(fields) <> 4 Then Exit Function
    If fields(0) <> "BTS" Then Exit Function
    ComponentMaterial = Left$(CStr(fields(2)), 2)
End Function

Public Sub BTS_AddItem()
    On Error GoTo Failed
    Application.EnableEvents = False
    If BTS_FormIsAssembly() Then BTS_SortDraft
    SaveCurrent False
Clean:
    Application.EnableEvents = True
    Exit Sub
Failed:
    MsgBox "BTS_AddItem: " & Err.Description, vbExclamation
    Resume Clean
End Sub

Private Sub RemoveCompositions(ByVal parent As String)
    Dim data As Worksheet, r As Long
    Set data = DB
    For r = EndRow(data, 6) To 2 Step -1
        If CStr(data.Cells(r, 6).Value) = parent Then data.Range("F" & r & ":I" & r).Delete Shift:=xlUp
    Next r
End Sub

Private Sub SaveCurrent(ByVal overwrite As Boolean)
    Dim sh As Worksheet, data As Worksheet, code As String, ar As Variant, r As Long, targetRow As Long, cmp As String, n As String, sameKind As Boolean
    Set sh = S(1)
    Set data = DB
    code = CodeFrom(sh)
    If Not ValidateForm(code, overwrite) Then Exit Sub
    ar = Split(code, ".")
    sameKind = IsAssemblyCode(code)
    If sameKind Then
        n = Trim$(CStr(sh.Range("C6").Value))
        If ar(1) = "SH" Then n = AutomaticName(code)
        If n = "" Then MsgBox BTS_T(47), vbExclamation: Exit Sub
        If overwrite Then
            targetRow = AssemblyRow(code)
            If targetRow = 0 Then MsgBox BTS_T(49), vbExclamation: Exit Sub
        Else
            targetRow = NextFree(data, 1)
        End If
        data.Cells(targetRow, 1).Value = code
        data.Cells(targetRow, 2).Value = n
        data.Cells(targetRow, 3).Value = ar(1)
        RemoveCompositions code
        For r = FORM_FIRST To FORM_LAST
            cmp = Trim$(CStr(sh.Cells(r, 2).Value))
            If cmp <> "" Then
                targetRow = NextFree(data, 6)
                data.Cells(targetRow, 6).Value = code
                data.Cells(targetRow, 7).Value = cmp
                data.Cells(targetRow, 8).Value = CLng(sh.Cells(r, 4).Value)
                data.Cells(targetRow, 9).Value = Trim$(CStr(sh.Cells(r, 5).Value))
            End If
        Next r
    Else
        If overwrite Then
            targetRow = PartRow(code)
            If targetRow = 0 Then MsgBox BTS_T(49), vbExclamation: Exit Sub
        Else
            targetRow = NextFree(S(2), 1)
        End If
        S(2).Cells(targetRow, 1).Value = code
        S(2).Cells(targetRow, 2).Value = CStr(sh.Range("C6").Value)
        S(2).Cells(targetRow, 3).Value = CStr(sh.Range("C7").Value)
    End If
    BTS_RebuildAll
    MsgBox BTS_T(37), vbInformation
End Sub

' The former load, save and propose actions are removed.

Private Sub BTS_ComponentChoices()
    ' Own parts and registered subassemblies use one selector; purchased items another.
    Dim data As Worksheet, sh As Worksheet, i As Long, p As Long, q As Long, code As String
    Set data = DB
    data.Range("AI2:AJ" & Application.Max(3000, EndRow(data, 35), EndRow(data, 36))).ClearContents
    p = 2
    q = 2
    Set sh = S(2)
    For i = 2 To EndRow(sh, 1)
        code = Trim$(CStr(sh.Cells(i, 1).Value))
        If code <> "" Then
            data.Cells(p, 35).Value = code
            p = p + 1
        End If
    Next i
    For i = 2 To EndRow(data, 1)
        code = Trim$(CStr(data.Cells(i, 1).Value))
        If code <> "" Then
            data.Cells(p, 35).Value = code
            p = p + 1
        End If
    Next i
    Set sh = S(7)
    For i = 2 To EndRow(sh, 1)
        code = Trim$(CStr(sh.Cells(i, 1).Value))
        If code <> "" Then
            data.Cells(q, 36).Value = code
            q = q + 1
        End If
    Next i
    If p > 3 Then data.Range("AI2:AI" & CStr(p - 1)).Sort Key1:=data.Range("AI2"), Header:=xlNo, Order1:=xlAscending
    If q > 3 Then data.Range("AJ2:AJ" & CStr(q - 1)).Sort Key1:=data.Range("AJ2"), Header:=xlNo, Order1:=xlAscending
    EnsureName "BTSDraftOwn", "'_BTS_DATA'!$AI$2:$AI$" & CStr(Application.Max(2, p - 1))
    EnsureName "BTSDraftStandard", "'_BTS_DATA'!$AJ$2:$AJ$" & CStr(Application.Max(2, q - 1))
    PutList S(1).Range("C10:D10"), "BTSDraftOwn"
    PutList S(1).Range("C11:D11"), "BTSDraftStandard"
End Sub

Private Function DraftRole(ByVal code As String) As String
    ' Only the designation and quantity are selected; infer the editable role from the registered name.
    Dim n As String
    n = ItemName(code)
    DraftRole = BTS_T(24)
    If StandardRow(code) > 0 Then Exit Function
    If InStr(1, n, BTS_T(22), vbTextCompare) > 0 Then DraftRole = BTS_T(22): Exit Function
    If InStr(1, n, BTS_T(23), vbTextCompare) > 0 Then DraftRole = BTS_T(23)
End Function

Public Sub BTS_AddDraftOwn()
    BTS_AddDraft 10, False
End Sub

Public Sub BTS_AddDraftStandard()
    BTS_AddDraft 11, True
End Sub

Private Sub BTS_AddDraft(ByVal selectorRow As Long, ByVal standardOnly As Boolean)
    ' Add a selected position or increase its existing quantity.
    Dim sh As Worksheet, code As String, q As Variant, total As Double, r As Long, dest As Long, existing As Long
    On Error GoTo Failed
    Set sh = S(1)
    If Not BTS_FormIsAssembly() Then Exit Sub
    code = Trim$(CStr(sh.Cells(selectorRow, 3).Value))
    q = sh.Cells(selectorRow, 5).Value
    If code = "" Or Not ExistsCode(code) Or Not ValidDraftQty(q) Then MsgBox BTS_T(86), vbExclamation: Exit Sub
    If standardOnly Then
        If StandardRow(code) = 0 Then MsgBox BTS_T(86), vbExclamation: Exit Sub
    Else
        If PartRow(code) = 0 And AssemblyRow(code) = 0 Then MsgBox BTS_T(86), vbExclamation: Exit Sub
    End If
    For r = FORM_FIRST To FORM_LAST
        If StrComp(Trim$(CStr(sh.Cells(r, 2).Value)), code, vbTextCompare) = 0 Then existing = r
        If dest = 0 And Trim$(CStr(sh.Cells(r, 2).Value)) = "" Then dest = r
    Next r
    If existing > 0 Then
        If Not ValidDraftQty(sh.Cells(existing, 4).Value) Then MsgBox BTS_T(86), vbExclamation: Exit Sub
        total = CDbl(sh.Cells(existing, 4).Value) + CDbl(q)
        If Not ValidDraftQty(total) Then MsgBox BTS_T(86), vbExclamation: Exit Sub
    ElseIf dest = 0 Then
        MsgBox BTS_T(87), vbExclamation
        Exit Sub
    End If
    Application.EnableEvents = False
    If existing > 0 Then
        sh.Cells(existing, 4).Value = CLng(total)
    Else
        sh.Cells(dest, 2).Value = code
        sh.Cells(dest, 3).Value = ItemName(code)
        sh.Cells(dest, 4).Value = CLng(q)
        sh.Cells(dest, 5).Value = DraftRole(code)
    End If
    BTS_SortDraft
    sh.Range("C" & selectorRow & ":D" & selectorRow).ClearContents
    sh.Cells(selectorRow, 5).Value = 1
Clean:
    Application.EnableEvents = True
    Exit Sub
Failed:
    MsgBox "BTS_AddDraft: " & Err.Description, vbExclamation
    Resume Clean
End Sub

Private Function NaturalKey(ByVal value As String) As String
    ' Pad numeric runs so 2, 8, 10 sort by numeric value rather than as text.
    Dim i As Long, ch As String, digits As String, result As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        Else
            If Len(digits) > 0 Then result = result & Right$(String$(14, "0") & digits, 14): digits = ""
            result = result & UCase$(ch)
        End If
    Next i
    If Len(digits) > 0 Then result = result & Right$(String$(14, "0") & digits, 14)
    NaturalKey = result
End Function

Private Function StandardKey(ByVal code As String) As String
    ' ESKD order for the agreed free-text standard-product designation convention.
    ' Foreign/international, interstate, national, company, industry, then unclassified.
    Dim category As String, family As String, nameKey As String, standard As String, p As Long, token As Variant, n As String, firstWord As String, rest As String
    category = "9"
    standard = ""
    For Each token In Array("DIN", "ISO", "EN", "ANSI", "ASME", "BS", BTS_T(89), BTS_T(90), BTS_T(91), BTS_T(92))
        p = InStr(1, code, CStr(token) & " ", vbTextCompare)
        If p > 0 Then
            rest = Trim$(Mid$(code, p + Len(CStr(token))))
            standard = CStr(token) & " " & Split(rest & " ", " ")(0)
            If CStr(token) = BTS_T(89) Then
                category = "3"
            ElseIf CStr(token) = BTS_T(90) Then
                category = "2"
            ElseIf CStr(token) = BTS_T(91) Then
                category = "4"
            ElseIf CStr(token) = BTS_T(92) Then
                category = "5"
            Else
                category = "1"
            End If
            Exit For
        End If
    Next token
    n = Trim$(ItemName(code))
    If n = "" Then n = code
    firstWord = Split(n & " ", " ")(0)
    nameKey = UCase$(firstWord)
    family = nameKey
    For Each token In Array(BTS_T(93), BTS_T(94), BTS_T(95), BTS_T(96), BTS_T(97))
        If StrComp(firstWord, CStr(token), vbTextCompare) = 0 Then
            family = BTS_T(98)
            Exit For
        End If
    Next token
    ' Use the standard identifier before the size; retain original text as the final tie breaker.
    StandardKey = category & "|" & family & "|" & nameKey & "|" & NaturalKey(standard) & "|" & NaturalKey(code)
End Function

Private Function SpecKey(ByVal code As String) As String
    Dim k As Long
    k = Kind(code)
    If k = 2 Then
        SpecKey = "2|" & StandardKey(code)
    Else
        SpecKey = CStr(k) & "|" & NaturalKey(code)
    End If
End Function

Private Sub BTS_SortDraft()
    ' Rebuild compact rows from the editable table. The order of each row's fields is preserved.
    Dim sh As Worksheet, rows() As Variant, tmp As Variant, n As Long, r As Long, i As Long, j As Long, c As Long, code As String, btnShape As Shape
    Set sh = S(1)
    For i = sh.Shapes.Count To 1 Step -1
        If Left$(sh.Shapes(i).Name, 9) = "draftDel_" Then sh.Shapes(i).Delete
    Next i
    ReDim rows(1 To FORM_LAST - FORM_FIRST + 1, 1 To 4)
    For r = FORM_FIRST To FORM_LAST
        code = Trim$(CStr(sh.Cells(r, 2).Value))
        If code <> "" Then
            n = n + 1
            rows(n, 1) = code
            rows(n, 2) = ItemName(code)
            rows(n, 3) = sh.Cells(r, 4).Value
            rows(n, 4) = sh.Cells(r, 5).Value
            If Trim$(CStr(rows(n, 4))) = "" Then rows(n, 4) = DraftRole(code)
        End If
    Next r
    For i = 1 To n - 1
        For j = i + 1 To n
            If StrComp(SpecKey(CStr(rows(i, 1))), SpecKey(CStr(rows(j, 1))), vbTextCompare) > 0 Then
                For c = 1 To 4
                    tmp = rows(i, c): rows(i, c) = rows(j, c): rows(j, c) = tmp
                Next c
            End If
        Next j
    Next i
    sh.Range("A" & FORM_FIRST & ":F" & FORM_LAST).ClearContents
    For i = 1 To n
        r = FORM_FIRST + i - 1
        sh.Cells(r, 1).Value = i
        For c = 1 To 4
            sh.Cells(r, c + 1).Value = rows(i, c)
        Next c
        If BTS_FormIsAssembly() Then
            MakeButton sh, "draftDel_" & r, BTS_T(81), "BTS_RemoveDraft", "F" & r, 65
            Set btnShape = sh.Shapes("draftDel_" & r)
            btnShape.Height = 22
            btnShape.Width = sh.Cells(r, 6).Width - 4
            btnShape.Top = sh.Cells(r, 6).Top + 3
            btnShape.Left = sh.Cells(r, 6).Left + 2
            btnShape.Placement = xlMoveAndSize
        End If
    Next i
End Sub

Public Sub BTS_RemoveDraft()
    ' Remove a row of the unsaved draft only. Never alter the catalogue.
    Dim sh As Worksheet, r As Long, callerName As Variant
    On Error GoTo Failed
    Set sh = S(1)
    If Not BTS_FormIsAssembly() Then Exit Sub
    callerName = Application.Caller
    If VarType(callerName) = vbString Then
        If Left$(CStr(callerName), 9) = "draftDel_" Then r = Val(Mid$(CStr(callerName), 10))
    Else
        r = ActiveCell.Row
    End If
    If r < FORM_FIRST Or r > FORM_LAST Then Exit Sub
    Application.EnableEvents = False
    sh.Range("A" & r & ":F" & r).ClearContents
    BTS_SortDraft
Clean:
    Application.EnableEvents = True
    Exit Sub
Failed:
    MsgBox "BTS_RemoveDraft: " & Err.Description, vbExclamation
    Resume Clean
End Sub

Private Sub SortChildren(ByRef pairs As Variant, ByVal qty As Long)
    Dim i As Long, j As Long, tmp As Variant
    For i = 1 To qty - 1
        For j = i + 1 To qty
            If StrComp(SpecKey(CStr(pairs(i, 1))), SpecKey(CStr(pairs(j, 1))), vbTextCompare) > 0 Then
                tmp = Array(pairs(i, 1), pairs(i, 2))
                pairs(i, 1) = pairs(j, 1): pairs(i, 2) = pairs(j, 2)
                pairs(j, 1) = tmp(0): pairs(j, 2) = tmp(1)
            End If
        Next j
    Next i
End Sub

Private Function Children(ByVal parent As String, ByRef countOut As Long) As Variant
    Dim data As Worksheet, i As Long, matrix() As Variant
    Set data = DB
    countOut = 0
    For i = 2 To EndRow(data, 6)
        If CStr(data.Cells(i, 6).Value) = parent Then countOut = countOut + 1
    Next i
    ReDim matrix(1 To Application.Max(countOut, 1), 1 To 2)
    countOut = 0
    For i = 2 To EndRow(data, 6)
        If CStr(data.Cells(i, 6).Value) = parent Then
            countOut = countOut + 1
            matrix(countOut, 1) = CStr(data.Cells(i, 7).Value)
            matrix(countOut, 2) = CLng(data.Cells(i, 8).Value)
        End If
    Next i
    If countOut > 1 Then SortChildren matrix, countOut
    Children = matrix
End Function

Private Sub CopyLink(ByVal code As String, ByVal dest As Range)
    Dim p As Long, linkAddress As String, innerAddress As String
    p = PartRow(code)
    If p > 0 Then
        If S(2).Cells(p, 1).Hyperlinks.Count > 0 Then
            linkAddress = S(2).Cells(p, 1).Hyperlinks(1).Address
            innerAddress = S(2).Cells(p, 1).Hyperlinks(1).SubAddress
        End If
    Else
        p = AssemblyRow(code)
        If p > 0 Then
            linkAddress = CStr(DB.Cells(p, 4).Value)
            innerAddress = CStr(DB.Cells(p, 5).Value)
        Else
            p = StandardRow(code)
            If p > 0 Then
                If S(7).Cells(p, 1).Hyperlinks.Count > 0 Then
                    linkAddress = S(7).Cells(p, 1).Hyperlinks(1).Address
                    innerAddress = S(7).Cells(p, 1).Hyperlinks(1).SubAddress
                End If
            End If
        End If
    End If
    If linkAddress <> "" Or innerAddress <> "" Then
        dest.Parent.Hyperlinks.Add Anchor:=dest, Address:=linkAddress, SubAddress:=innerAddress, TextToDisplay:=code
    End If
End Sub

Private Sub SyncTreeLinks()
    Dim i As Long, sh As Worksheet, r As Long, d As Long, code As String, p As Long
    For i = 3 To 6
        Set sh = S(i)
        For r = 2 To EndRow(sh, 11)
            For d = 1 To MAX_DEPTH
                If sh.Cells(r, d).Hyperlinks.Count > 0 Then
                    code = CStr(sh.Cells(r, d).Value)
                    p = AssemblyRow(code)
                    If p > 0 Then
                        DB.Cells(p, 4).Value = sh.Cells(r, d).Hyperlinks(1).Address
                        DB.Cells(p, 5).Value = sh.Cells(r, d).Hyperlinks(1).SubAddress
                    Else
                        p = PartRow(code)
                        If p > 0 Then
                            If S(2).Cells(p, 1).Hyperlinks.Count = 0 Then
                                S(2).Hyperlinks.Add Anchor:=S(2).Cells(p, 1), Address:=sh.Cells(r, d).Hyperlinks(1).Address, SubAddress:=sh.Cells(r, d).Hyperlinks(1).SubAddress, TextToDisplay:=code
                            End If
                        End If
                    End If
                End If
            Next d
        Next r
    Next i
End Sub

Private Function IsExcluded(ByVal root As String, ByVal path As String) As Boolean
    Dim data As Worksheet, r As Long
    Set data = DB
    For r = 2 To EndRow(data, 14)
        If CStr(data.Cells(r, 14).Value) = root And CStr(data.Cells(r, 15).Value) = path Then IsExcluded = True: Exit Function
    Next r
End Function

Private Sub DrawNode(ByVal sh As Worksheet, ByVal code As String, ByVal parent As String, ByVal root As String, ByVal path As String, ByVal depth As Long, ByVal multiplier As Double, ByRef rowOut As Long, ByVal orderMode As Boolean)
    Dim startRow As Long, rowN As Long, ct As Long, i As Long, x As Variant, child As String, childPath As String
    If depth > MAX_DEPTH Then Err.Raise vbObjectError + 200, , "Assembly depth exceeds 6 columns"
    If orderMode And IsExcluded(root, path) Then Exit Sub
    startRow = rowOut
    sh.Cells(startRow, depth).Value = code
    sh.Cells(startRow, 7).Value = ItemName(code)
    sh.Cells(startRow, 8).Value = MaterialMark(code)
    If parent <> "" Then sh.Cells(startRow, 9).Value = multiplier
    If orderMode Then sh.Cells(startRow, 9).Value = multiplier
    sh.Cells(startRow, 10).Value = parent
    sh.Cells(startRow, 11).Value = path
    CopyLink code, sh.Cells(startRow, depth)
    If IsAssembly(code) Then
        With sh.Range(sh.Cells(startRow, depth), sh.Cells(startRow, 9))
            .Interior.Color = RGB(236, 240, 242)
        End With
    End If
    rowOut = rowOut + 1
    If Not IsAssembly(code) Then Exit Sub
    x = Children(code, ct)
    For i = 1 To ct
        child = CStr(x(i, 1))
        childPath = path & "/" & child
        DrawNode sh, child, code, root, childPath, depth + 1, multiplier * CDbl(x(i, 2)), rowOut, orderMode
    Next i
    If rowOut - startRow > 1 Then sh.Rows((startRow + 1) & ":" & (rowOut - 1)).Group
End Sub

Private Sub ResetTree(ByVal sh As Worksheet, ByVal firstRow As Long)
    Dim last As Long
    last = Application.Max(EndRow(sh, 11), firstRow + 1)
    On Error Resume Next
    sh.Rows(firstRow & ":" & last).ClearOutline
    sh.Range("A" & firstRow & ":K" & last).Hyperlinks.Delete
    sh.Range("A" & firstRow & ":K" & last).ClearContents
    sh.Range("A" & firstRow & ":I" & last).Interior.Pattern = xlNone
    On Error GoTo 0
End Sub

Private Function FamilySortKey(ByVal code As String) As String
    ' Diameter + angle precede material when comparing finished components.
    Dim ar As Variant
    ar = Split(code, ".")
    If UBound(ar) = 4 Then
        FamilySortKey = CStr(ar(3)) & "." & CStr(ar(2)) & "." & CStr(ar(4))
    Else
        FamilySortKey = code
    End If
End Function

Private Sub DrawFamily(ByVal id As Long)
    Dim sh As Worksheet, data As Worksheet, n As Long, i As Long, j As Long, total As Long, code As String, temp As String, arr() As String, rowN As Long
    Set sh = S(id)
    Set data = DB
    ResetTree sh, 2
    total = 0
    For i = 2 To EndRow(data, 1)
        If CStr(data.Cells(i, 3).Value) = Choose(id - 2, "SH", "KR", "DU", "DT") Then total = total + 1
    Next i
    If total = 0 Then Exit Sub
    ReDim arr(1 To total)
    n = 0
    For i = 2 To EndRow(data, 1)
        If CStr(data.Cells(i, 3).Value) = Choose(id - 2, "SH", "KR", "DU", "DT") Then
            n = n + 1
            arr(n) = CStr(data.Cells(i, 1).Value)
        End If
    Next i
    For i = 1 To total - 1
        For j = i + 1 To total
            If StrComp(FamilySortKey(arr(i)), FamilySortKey(arr(j)), vbTextCompare) > 0 Then temp = arr(i): arr(i) = arr(j): arr(j) = temp
        Next j
    Next i
    rowN = 2
    sh.Outline.SummaryRow = xlAbove
    For i = 1 To total
        code = arr(i)
        DrawNode sh, code, "", code, code, 1, 1, rowN, False
    Next i
    sh.Rows("2:" & rowN).RowHeight = 20
End Sub

Public Sub BTS_RebuildAll()
    Dim i As Long
    On Error GoTo Failed
    Application.ScreenUpdating = False
    If ThisWorkbook.Worksheets("_BTS_DATA").Visible <> xlSheetVeryHidden Then ThisWorkbook.Worksheets("_BTS_DATA").Visible = xlSheetVeryHidden
    SyncTreeLinks
    BTS_ComponentChoices
    For i = 3 To 6
        DrawFamily i
    Next i
    BTS_RebuildOrder
    BTS_UpdateOrderLists
Clean:
    Application.ScreenUpdating = True
    Exit Sub
Failed:
    MsgBox "BTS_RebuildAll: " & Err.Description, vbExclamation
    Resume Clean
End Sub

Private Sub RemoveCompEdges(ByVal parent As String, ByVal child As String)
    Dim data As Worksheet, r As Long
    Set data = DB
    For r = EndRow(data, 6) To 2 Step -1
        If CStr(data.Cells(r, 6).Value) = parent And CStr(data.Cells(r, 7).Value) = child Then data.Range("F" & r & ":I" & r).Delete Shift:=xlUp
    Next r
End Sub

Public Sub BTS_DeleteSelected()
    Dim sh As Worksheet, idx As Long, r As Long, code As String, parent As String, path As String, root As String, useText As String, p As Long, rr As Long
    Set sh = ActiveSheet
    r = ActiveCell.Row
    idx = 0
    If sh.Name = BTS_T(2) Then idx = 2
    If sh.Name = BTS_T(7) Then idx = 7
    If sh.Name = BTS_T(8) Then idx = 8
    If r < 2 Then Exit Sub
    If idx = 2 Or idx = 7 Then
        code = CStr(sh.Cells(r, 1).Value)
        If code = "" Then Exit Sub
        For rr = 2 To EndRow(DB, 1)
            If CStr(DB.Cells(rr, 1).Value) <> "" Then
                If HasPath(CStr(DB.Cells(rr, 1).Value), code, 0) Then
                    useText = useText & vbCrLf & CStr(DB.Cells(rr, 1).Value)
                End If
            End If
        Next rr
        If MsgBox(BTS_T(52) & code & vbCrLf & useText, vbYesNo + vbExclamation) <> vbYes Then Exit Sub
        For rr = EndRow(DB, 6) To 2 Step -1
            If CStr(DB.Cells(rr, 7).Value) = code Then DB.Range("F" & rr & ":I" & rr).Delete Shift:=xlUp
        Next rr
        sh.Rows(r).Delete
        BTS_RebuildAll
        Exit Sub
    End If
    If idx = 8 Then
        path = CStr(sh.Cells(r, 11).Value)
        If path = "" Then Exit Sub
        root = Split(path, "/")(0)
        If MsgBox(BTS_T(53), vbYesNo + vbQuestion) <> vbYes Then Exit Sub
        If path = root Then
            For rr = EndRow(DB, 11) To 2 Step -1
                If CStr(DB.Cells(rr, 11).Value) = root Then DB.Range("K" & rr & ":L" & rr).Delete Shift:=xlUp
            Next rr
            For rr = EndRow(DB, 14) To 2 Step -1
                If CStr(DB.Cells(rr, 14).Value) = root Then DB.Range("N" & rr & ":O" & rr).Delete Shift:=xlUp
            Next rr
        Else
            rr = NextFree(DB, 14)
            DB.Cells(rr, 14).Value = root
            DB.Cells(rr, 15).Value = path
        End If
        BTS_RebuildOrder
        Exit Sub
    End If
    If sh.Name = BTS_T(3) Or sh.Name = BTS_T(4) Or sh.Name = BTS_T(5) Or sh.Name = BTS_T(6) Then
        code = CStr(sh.Cells(r, 11).Value)
        parent = CStr(sh.Cells(r, 10).Value)
        If code = "" Then Exit Sub
        If parent <> "" Then
            Dim arr As Variant
            arr = Split(code, "/")
            code = CStr(arr(UBound(arr)))
            If MsgBox(BTS_T(54) & code & vbCrLf & BTS_T(55) & parent, vbYesNo + vbQuestion) <> vbYes Then Exit Sub
            RemoveCompEdges parent, code
        Else
            code = CStr(sh.Cells(r, 1).Value)
            If MsgBox(BTS_T(52) & code, vbYesNo + vbExclamation) <> vbYes Then Exit Sub
            rr = AssemblyRow(code)
            If rr > 0 Then DB.Range("A" & rr & ":E" & rr).Delete Shift:=xlUp
            RemoveCompositions code
            For rr = EndRow(DB, 11) To 2 Step -1
                If CStr(DB.Cells(rr, 11).Value) = code Then DB.Range("K" & rr & ":L" & rr).Delete Shift:=xlUp
            Next rr
            For rr = EndRow(DB, 14) To 2 Step -1
                If CStr(DB.Cells(rr, 14).Value) = code Then DB.Range("N" & rr & ":O" & rr).Delete Shift:=xlUp
            Next rr
            For rr = EndRow(DB, 6) To 2 Step -1
                If CStr(DB.Cells(rr, 7).Value) = code Then DB.Range("F" & rr & ":I" & rr).Delete Shift:=xlUp
            Next rr
        End If
        BTS_RebuildAll
    End If
End Sub

Private Sub BTS_RebuildOrder()
    ' Sort complete order roots, never reorder individual rows inside a BOM.
    Dim sh As Worksheet, data As Worksheet
    Dim i As Long, j As Long, n As Long, rowN As Long, code As String, swap As String
    Dim roots() As String
    Set sh = S(8)
    Set data = DB
    ResetTree sh, 8
    sh.Range("L8:L10000").ClearContents
    n = 0
    For i = 2 To EndRow(data, 11)
        If CStr(data.Cells(i, 11).Value) <> "" Then n = n + 1
    Next i
    If n > 0 Then
        ReDim roots(1 To n)
        n = 0
        For i = 2 To EndRow(data, 11)
            code = CStr(data.Cells(i, 11).Value)
            If code <> "" Then
                n = n + 1
                roots(n) = code
            End If
        Next i
        For i = 1 To n - 1
            For j = i + 1 To n
                If StrComp(roots(i), roots(j), vbTextCompare) > 0 Then
                    swap = roots(i)
                    roots(i) = roots(j)
                    roots(j) = swap
                End If
            Next j
        Next i
    End If
    rowN = 8
    sh.Outline.SummaryRow = xlAbove
    For i = 1 To n
        code = roots(i)
        j = Locate(data, 11, code)
        If j > 0 Then DrawNode sh, code, "", code, code, 1, CDbl(data.Cells(j, 12).Value), rowN, True
    Next i
    If rowN > 8 Then
        sh.Range("L8:L" & CStr(rowN - 1)).Value = BTS_T(25)
        sh.Range("L8:L" & CStr(rowN - 1)).Font.Color = RGB(95, 39, 39)
        sh.Range("L8:L" & CStr(rowN - 1)).HorizontalAlignment = xlCenter
    End If
End Sub

Public Sub BTS_AddToOrder()
    Dim code As String, qty As Double, r As Long, sh As Worksheet
    Set sh = S(8)
    code = CodeFrom(sh)
    If Not ExistsCode(code) Then MsgBox BTS_T(41), vbExclamation: Exit Sub
    If Not IsNumeric(sh.Range("H2").Value) Then MsgBox BTS_T(56), vbExclamation: Exit Sub
    qty = CDbl(sh.Range("H2").Value)
    If qty < 1 Or Fix(qty) <> qty Then MsgBox BTS_T(56), vbExclamation: Exit Sub
    r = Locate(DB, 11, code)
    If r = 0 Then r = NextFree(DB, 11): DB.Cells(r, 11).Value = code
    DB.Cells(r, 12).Value = Val(DB.Cells(r, 12).Value) + qty
    BTS_RebuildOrder
End Sub

Public Sub BTS_ClearOrder()
    If MsgBox(BTS_T(57), vbYesNo + vbQuestion) <> vbYes Then Exit Sub
    DB.Range("K2:L2000").ClearContents
    DB.Range("N2:O3000").ClearContents
    BTS_RebuildOrder
End Sub

Private Sub CountLeaves(ByVal root As String, ByVal code As String, ByVal path As String, ByVal n As Double, ByVal totals As Object, ByVal depth As Long)
    Dim i As Long, data As Worksheet
    If depth > 12 Then Err.Raise vbObjectError + 201, , "Cyclic or overdeep BOM"
    If IsExcluded(root, path) Then Exit Sub
    If Not IsAssembly(code) Then
        If totals.Exists(code) Then
            totals(code) = totals(code) + n
        Else
            totals.Add code, n
        End If
        Exit Sub
    End If
    Set data = DB
    Dim foundChild As Boolean
    For i = 2 To EndRow(data, 6)
        If CStr(data.Cells(i, 6).Value) = code Then
            foundChild = True
            CountLeaves root, CStr(data.Cells(i, 7).Value), path & "/" & CStr(data.Cells(i, 7).Value), n * CDbl(data.Cells(i, 8).Value), totals, depth + 1
        End If
    Next i
    If Not foundChild Then Err.Raise vbObjectError + 202, , "No components defined: " & code
End Sub

Public Sub BTS_ExportOrder()
    Dim totals As Object, data As Worksheet, output As Workbook, sh As Worksheet, i As Long, r As Long, k As Variant, fn As String, last As Long
    On Error GoTo Failed
    If ThisWorkbook.Path = "" Then MsgBox BTS_T(58), vbExclamation: Exit Sub
    Set totals = CreateObject("Scripting.Dictionary")
    Set data = DB
    For i = 2 To EndRow(data, 11)
        If CStr(data.Cells(i, 11).Value) <> "" Then CountLeaves CStr(data.Cells(i, 11).Value), CStr(data.Cells(i, 11).Value), CStr(data.Cells(i, 11).Value), CDbl(data.Cells(i, 12).Value), totals, 0
    Next i
    If totals.Count = 0 Then MsgBox BTS_T(59), vbExclamation: Exit Sub
    Application.ScreenUpdating = False
    Set output = Application.Workbooks.Add(xlWBATWorksheet)
    Set sh = output.Worksheets(1)
    sh.Name = "BTS_Order"
    sh.Range("A1:E1").Value = Array(BTS_T(60), BTS_T(10), BTS_T(11), BTS_T(12), BTS_T(13))
    r = 2
    For Each k In totals.Keys
        sh.Cells(r, 1).Value = IIf(PartRow(CStr(k)) > 0, BTS_T(38), BTS_T(70))
        sh.Cells(r, 2).Value = CStr(k)
        sh.Cells(r, 3).Value = ItemName(CStr(k))
        sh.Cells(r, 4).Value = MaterialMark(CStr(k))
        sh.Cells(r, 5).Value = totals(k)
        r = r + 1
    Next k
    last = r - 1
    sh.Range("A1:E" & last).Sort Key1:=sh.Range("A2"), Order1:=xlAscending, Key2:=sh.Range("B2"), Order2:=xlAscending, Header:=xlYes
    sh.Columns("A:E").AutoFit
    sh.Range("A1:E1").Interior.Color = RGB(227, 232, 236)
    sh.Range("A1:E" & last).Borders.LineStyle = xlContinuous
    fn = ThisWorkbook.Path & Application.PathSeparator & "BTS_Order_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsx"
    output.SaveAs Filename:=fn, FileFormat:=xlOpenXMLWorkbook
    output.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox BTS_T(71) & vbCrLf & fn, vbInformation
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    MsgBox "BTS_ExportOrder: " & Err.Description, vbExclamation
End Sub

Public Sub BTS_UpdateOrderLists()
    Dim sh As Worksheet, data As Worksheet, entries As Object, key As Variant, p As Long, c As Long, i As Long, cnt As Long, code As String, arr As Variant, piece(1 To 6) As String, prevOK As Boolean, colName As String, nm As String
    Set sh = S(8)
    Set data = DB
    For c = 1 To 6
        Set entries = CreateObject("Scripting.Dictionary")
        For i = 2 To EndRow(data, 1)
            code = CStr(data.Cells(i, 1).Value)
            If code <> "" Then Call AddOption(code, sh, c, entries)
        Next i
        For i = 2 To EndRow(S(2), 1)
            code = CStr(S(2).Cells(i, 1).Value)
            If code <> "" Then Call AddOption(code, sh, c, entries)
        Next i
        data.Range(data.Cells(2, 26 + c), data.Cells(3000, 26 + c)).ClearContents
        p = 2
        For Each key In entries.Keys
            data.Cells(p, 26 + c).Value = key
            p = p + 1
        Next key
        If p = 2 Then data.Cells(p, 26 + c).Value = ""
        nm = "BTSOrder" & CStr(c)
        EnsureName nm, "'_BTS_DATA'!" & data.Range(data.Cells(2, 26 + c), data.Cells(Application.Max(p - 1, 2), 26 + c)).Address
        PutList sh.Cells(2, c), nm
    Next c
End Sub

Private Sub AddOption(ByVal code As String, ByVal formSh As Worksheet, ByVal targetIndex As Long, ByVal dict As Object)
    Dim ar As Variant, v(1 To 6) As String, x As Long, raw As String, ok As Boolean
    ar = Split(code, ".")
    If UBound(ar) <> 4 Then Exit Sub
    v(1) = ar(1)
    v(2) = Left$(ar(2), 2)
    If Len(ar(2)) = 4 Then v(3) = Right$(ar(2), 2)
    v(4) = Left$(ar(3), 2)
    v(5) = Right$(ar(3), 2)
    v(6) = ar(4)
    For x = 1 To targetIndex - 1
        raw = Trim$(CStr(formSh.Cells(2, x).Value))
        If raw <> "" And raw <> v(x) Then Exit Sub
    Next x
    If Not dict.Exists(v(targetIndex)) Then dict.Add v(targetIndex), True
End Sub

Public Sub BTS_Changed(ByVal Sh As Object, ByVal Target As Range)
    Dim raw As String, r As Long
    On Error GoTo Failed
    Application.EnableEvents = False
    If Sh.Name = BTS_T(8) Then
        If Not Intersect(Target, Sh.Range("A2:F2")) Is Nothing Then BTS_UpdateOrderLists
    ElseIf Sh.Name = BTS_T(7) Then
        If Not Intersect(Target, Sh.Range("A2:B2000")) Is Nothing Then BTS_ComponentChoices
    ElseIf Sh.Name = BTS_T(9) Then
        If Not Intersect(Target, Sh.Range("A17:B26")) Is Nothing Then
            If Not BTS_FormIsAssembly() Then BTS_MarkChoices
        End If
    ElseIf Sh.Name = BTS_T(1) Then
        If Not Intersect(Target, Sh.Range("C1")) Is Nothing Then
            BTS_FormMode True
        ElseIf Not Intersect(Target, Sh.Range("B4:G4")) Is Nothing Then
            If Not Intersect(Target, Sh.Range("B4")) Is Nothing Then Sh.Range("C6:G6").ClearContents
            If Not Intersect(Target, Sh.Range("C4")) Is Nothing Then
                If Not BTS_FormIsAssembly() Then BTS_MarkChoices
            End If
            If Not Intersect(Target, Sh.Range("G4")) Is Nothing Then
                raw = Trim$(CStr(Sh.Range("G4").Value2))
                If Len(raw) > 0 And Len(raw) <= 3 And Not raw Like "*[!0-9]*" Then
                    Sh.Range("G4").Value2 = Format$(CLng(raw), "000")
                    If (CLng(raw) Mod 10 = 0) <> BTS_FormIsAssembly() Then MsgBox BTS_T(77), vbExclamation
                ElseIf raw <> "" Then
                    MsgBox BTS_T(77), vbExclamation
                End If
            End If
        End If
        If Not Intersect(Target, Sh.Range("C1,B4:G4")) Is Nothing Then BTS_RefreshPreview
        If Not Intersect(Target, Sh.Range("B15:B54")) Is Nothing Then
            For r = FORM_FIRST To FORM_LAST
                If Not Intersect(Target, Sh.Cells(r, 2)) Is Nothing Then
                    Sh.Cells(r, 3).Value = ItemName(CStr(Sh.Cells(r, 2).Value))
                    Sh.Cells(r, 5).Value = DraftRole(CStr(Sh.Cells(r, 2).Value))
                End If
            Next r
            BTS_SortDraft
        End If
    End If
Clean:
    Application.EnableEvents = True
    Exit Sub
Failed:
    MsgBox "BTS_Changed: " & Err.Description, vbExclamation
    Resume Clean
End Sub
