' BTS-Catalog VBA main module. One record per item and reversible tree views.
Option Explicit

Private Const MAX_DEPTH As Long = 6
Private Const FORM_FIRST As Long = 9
Private Const FORM_LAST As Long = 48

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
    Dim sh As Worksheet
    Set sh = S(1)
    If CStr(sh.Range("A1").Value) = "" Then LayoutHeader sh, Array(BTS_T(14), BTS_T(15), BTS_T(16), BTS_T(17), BTS_T(18), BTS_T(19), BTS_T(10))
    sh.Columns("A:F").ColumnWidth = 17
    sh.Columns("G").ColumnWidth = 43
    sh.Columns("H").ColumnWidth = 22
    sh.Columns("I").ColumnWidth = 27
    sh.Range("A4").Value = BTS_T(11)
    sh.Range("A5").Value = BTS_T(12)
    sh.Range("A7").Value = BTS_T(20)
    sh.Range("A8:D8").Value = Array(BTS_T(10), BTS_T(11), BTS_T(13), BTS_T(21))
    sh.Range("A8:D8").Font.Bold = True
    sh.Range("A8:D8").Interior.Color = RGB(227, 232, 236)
    sh.Columns("A").ColumnWidth = 40
    sh.Columns("B").ColumnWidth = 42
    sh.Columns("C").ColumnWidth = 15
    sh.Columns("D").ColumnWidth = 17
    sh.Columns("E:F").ColumnWidth = 18
    PutList sh.Range("A2"), "BTSGroups"
    PutList sh.Range("B2:C2"), "BTSMats"
    PutList sh.Range("D2"), "BTSDiam"
    PutList sh.Range("E2"), "BTSAngles"
    PutList sh.Range("B5"), "BTSMarks"
    EnsureName "BTSRoles", "'_BTS_DATA'!$Z$1:$Z$3"
    DB.Range("Z1:Z3").Value = Application.Transpose(Array(BTS_T(22), BTS_T(23), BTS_T(24)))
    PutList sh.Range("D9:D48"), "BTSRoles"
    MakeButton sh, "addItem", BTS_T(26), "BTS_AddItem", "I2", 145
    MakeButton sh, "saveItem", BTS_T(27), "BTS_SaveItem", "I4", 145
    MakeButton sh, "loadItem", BTS_T(67), "BTS_LoadItem", "I6", 145
    MakeButton sh, "propose", BTS_T(68), "BTS_ProposeNumber", "I8", 145
    sh.Range("G2").Formula = "=IF(OR(A2="""",B2="""",D2="""",E2="""",F2=""""),"""",""BTS.""&A2&"".""&B2&IF(MOD(VALUE(F2),10)=0,C2,"""")&"".""&TEXT(VALUE(D2),""00"")&TEXT(VALUE(E2),""00"")&"".""&TEXT(VALUE(F2),""000""))"
    BTS_ComponentChoices
End Sub

Private Sub BTS_OrderInit()
    Dim sh As Worksheet
    Set sh = S(8)
    If CStr(sh.Range("A1").Value) = "" Then LayoutHeader sh, Array(BTS_T(14), BTS_T(15), BTS_T(16), BTS_T(17), BTS_T(18), BTS_T(19), BTS_T(10), BTS_T(13))
    sh.Columns("A:F").ColumnWidth = 16
    sh.Columns("G").ColumnWidth = 40
    sh.Columns("H").ColumnWidth = 17
    sh.Range("G2").Formula = S(1).Range("G2").Formula
    If IsEmpty(sh.Range("H2").Value) Then sh.Range("H2").Value = 1
    sh.Range("A7:I7").Value = Array("I", "II", "III", "IV", "V", "VI", BTS_T(11), BTS_T(12), BTS_T(13))
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
    Dim a As String, b As String, c As String, d As String, e As String, f As String
    a = Trim$(CStr(sh.Range("A2").Value))
    b = Trim$(CStr(sh.Range("B2").Value))
    c = Trim$(CStr(sh.Range("C2").Value))
    d = Trim$(CStr(sh.Range("D2").Value))
    e = Trim$(CStr(sh.Range("E2").Value))
    f = Trim$(CStr(sh.Range("F2").Value))
    If a = "" Or b = "" Or d = "" Or e = "" Or f = "" Then Exit Function
    If Not IsNumeric(d) Or Not IsNumeric(e) Or Not IsNumeric(f) Then Exit Function
    If CLng(f) < 0 Or CLng(f) > 999 Or CLng(d) < 0 Or CLng(d) > 99 Then Exit Function
    If CLng(f) Mod 10 = 0 Then
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
    For rowN = 17 To 26
        If StrComp(CStr(sh.Cells(rowN, 2).Value), markName, vbTextCompare) = 0 Then
            ValidMark = StrComp(CStr(sh.Cells(rowN, 1).Value), matCode, vbTextCompare) = 0
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

Private Function ValidateForm(ByVal code As String, ByVal overwrite As Boolean) As Boolean
    Dim sh As Worksheet, r As Long, cmp As String, q As Variant, nBase As Long, nHead As Long, ar As Variant, baseMat As String, headMat As String, kindAssembly As Boolean, seen As Object
    Set sh = S(1)
    If code = "" Then MsgBox BTS_T(40), vbExclamation: Exit Function
    If overwrite Then
        If Not ExistsCode(code) Then MsgBox BTS_T(41), vbExclamation: Exit Function
    Else
        If ExistsCode(code) Then MsgBox BTS_T(35), vbExclamation: Exit Function
    End If
    kindAssembly = IsAssemblyCode(code)
    ar = Split(code, ".")
    If kindAssembly Then
        Set seen = CreateObject("Scripting.Dictionary")
        If Len(ar(2)) <> 4 Then MsgBox BTS_T(42), vbExclamation: Exit Function
        baseMat = Left$(ar(2), 2)
        headMat = Right$(ar(2), 2)
        For r = FORM_FIRST To FORM_LAST
            cmp = Trim$(CStr(sh.Cells(r, 1).Value))
            If cmp <> "" Then
                If seen.Exists(cmp) Then MsgBox BTS_T(43), vbExclamation: Exit Function
                seen.Add cmp, True
                q = sh.Cells(r, 3).Value
                If Not ExistsCode(cmp) Or Not IsNumeric(q) Or Val(q) <= 0 Or Fix(Val(q)) <> Val(q) Then MsgBox BTS_T(43), vbExclamation: Exit Function
                If cmp = code Or HasPath(cmp, code, 0) Then MsgBox BTS_T(44), vbExclamation: Exit Function
                If CStr(sh.Cells(r, 4).Value) = BTS_T(22) Then
                    nBase = nBase + 1
                    If Mid$(Split(cmp, ".")(2), 1, 2) <> baseMat Then MsgBox BTS_T(45), vbExclamation: Exit Function
                ElseIf CStr(sh.Cells(r, 4).Value) = BTS_T(23) Then
                    nHead = nHead + 1
                    If Mid$(Split(cmp, ".")(2), 1, 2) <> headMat Then MsgBox BTS_T(45), vbExclamation: Exit Function
                End If
            End If
        Next r
        If ar(1) = "SH" Then
            If nBase <> 1 Or nHead <> 1 Then MsgBox BTS_T(46), vbExclamation: Exit Function
        End If
    Else
        If Trim$(CStr(sh.Range("B4").Value)) = "" Then MsgBox BTS_T(47), vbExclamation: Exit Function
        If Not ValidMark(CStr(sh.Range("B5").Value), CStr(ar(2))) Then MsgBox BTS_T(48), vbExclamation: Exit Function
    End If
    ValidateForm = True
End Function

Public Sub BTS_AddItem()
    SaveCurrent False
End Sub

Public Sub BTS_SaveItem()
    SaveCurrent True
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
        n = AutomaticName(code)
        If n = "" Then n = Trim$(CStr(sh.Range("B4").Value))
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
            cmp = Trim$(CStr(sh.Cells(r, 1).Value))
            If cmp <> "" Then
                targetRow = NextFree(data, 6)
                data.Cells(targetRow, 6).Value = code
                data.Cells(targetRow, 7).Value = cmp
                data.Cells(targetRow, 8).Value = CLng(sh.Cells(r, 3).Value)
                data.Cells(targetRow, 9).Value = CStr(sh.Cells(r, 4).Value)
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
        S(2).Cells(targetRow, 2).Value = CStr(sh.Range("B4").Value)
        S(2).Cells(targetRow, 3).Value = CStr(sh.Range("B5").Value)
    End If
    BTS_RebuildAll
    MsgBox BTS_T(37), vbInformation
End Sub

Public Sub BTS_LoadItem()
    Dim code As String, sh As Worksheet, data As Worksheet, p As Long, r As Long, atRow As Long, ar As Variant
    Set sh = S(1)
    Set data = DB
    code = CodeFrom(sh)
    If Not ExistsCode(code) Then MsgBox BTS_T(41), vbExclamation: Exit Sub
    ar = Split(code, ".")
    sh.Range("B4:B5").ClearContents
    sh.Range("A9:D48").ClearContents
    If IsAssembly(code) Then
        p = AssemblyRow(code)
        sh.Range("B4").Value = data.Cells(p, 2).Value
        atRow = FORM_FIRST
        For r = 2 To EndRow(data, 6)
            If CStr(data.Cells(r, 6).Value) = code Then
                If atRow > FORM_LAST Then Exit For
                sh.Cells(atRow, 1).Value = data.Cells(r, 7).Value
                sh.Cells(atRow, 2).Value = ItemName(CStr(data.Cells(r, 7).Value))
                sh.Cells(atRow, 3).Value = data.Cells(r, 8).Value
                sh.Cells(atRow, 4).Value = data.Cells(r, 9).Value
                atRow = atRow + 1
            End If
        Next r
    Else
        p = PartRow(code)
        sh.Range("B4").Value = S(2).Cells(p, 2).Value
        sh.Range("B5").Value = S(2).Cells(p, 3).Value
    End If
End Sub

Public Sub BTS_ProposeNumber()
    Dim sh As Worksheet, k As Long, code As String, unitType As Boolean
    Set sh = S(1)
    unitType = (MsgBox(BTS_T(50), vbYesNo + vbQuestion) = vbYes)
    For k = 0 To 999
        If (unitType And k Mod 10 = 0) Or (Not unitType And k Mod 10 <> 0) Then
            sh.Range("F2").Value = Format$(k, "000")
            code = CodeFrom(sh)
            If code <> "" And Not ExistsCode(code) Then Exit Sub
        End If
    Next k
    MsgBox BTS_T(51), vbExclamation
End Sub

Private Sub BTS_ComponentChoices()
    Dim data As Worksheet, sh As Worksheet, i As Long, p As Long
    Set data = DB
    data.Range("U2:U3000").ClearContents
    p = 2
    Set sh = S(2)
    For i = 2 To EndRow(sh, 1)
        If CStr(sh.Cells(i, 1).Value) <> "" Then data.Cells(p, 21).Value = CStr(sh.Cells(i, 1).Value): p = p + 1
    Next i
    Set sh = S(7)
    For i = 2 To EndRow(sh, 1)
        If CStr(sh.Cells(i, 1).Value) <> "" Then data.Cells(p, 21).Value = CStr(sh.Cells(i, 1).Value): p = p + 1
    Next i
    For i = 2 To EndRow(data, 1)
        If CStr(data.Cells(i, 1).Value) <> "" Then data.Cells(p, 21).Value = CStr(data.Cells(i, 1).Value): p = p + 1
    Next i
    If p = 2 Then data.Range("U2").Value = ""
    EnsureName "BTSComponents", "'_BTS_DATA'!$U$2:$U$" & CStr(Application.Max(2, p - 1))
    PutList S(1).Range("A9:A48"), "BTSComponents"
End Sub

Private Sub SortChildren(ByRef pairs As Variant, ByVal qty As Long)
    Dim i As Long, j As Long, tmp As Variant
    For i = 1 To qty - 1
        For j = i + 1 To qty
            If Kind(CStr(pairs(i, 1))) > Kind(CStr(pairs(j, 1))) Or (Kind(CStr(pairs(i, 1))) = Kind(CStr(pairs(j, 1))) And StrComp(CStr(pairs(i, 1)), CStr(pairs(j, 1)), vbTextCompare) > 0) Then
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
            If StrComp(arr(i), arr(j), vbTextCompare) > 0 Then temp = arr(i): arr(i) = arr(j): arr(j) = temp
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
        For rr = 2 To EndRow(DB, 6)
            If CStr(DB.Cells(rr, 7).Value) = code Then useText = useText & vbCrLf & CStr(DB.Cells(rr, 6).Value)
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
    Dim sh As Worksheet, data As Worksheet, i As Long, rowN As Long, code As String
    Set sh = S(8)
    Set data = DB
    ResetTree sh, 8
    rowN = 8
    sh.Outline.SummaryRow = xlAbove
    For i = 2 To EndRow(data, 11)
        code = CStr(data.Cells(i, 11).Value)
        If code <> "" Then DrawNode sh, code, "", code, code, 1, CDbl(data.Cells(i, 12).Value), rowN, True
    Next i
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
    For i = 2 To EndRow(data, 6)
        If CStr(data.Cells(i, 6).Value) = code Then CountLeaves root, CStr(data.Cells(i, 7).Value), path & "/" & CStr(data.Cells(i, 7).Value), n * CDbl(data.Cells(i, 8).Value), totals, depth + 1
    Next i
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
    If Sh.Name = BTS_T(8) Then
        If Not Intersect(Target, Sh.Range("A2:F2")) Is Nothing Then
            On Error GoTo Clean
            Application.EnableEvents = False
            BTS_UpdateOrderLists
        End If
    ElseIf Sh.Name = BTS_T(7) Then
        If Not Intersect(Target, Sh.Range("A2:B2000")) Is Nothing Then
            On Error GoTo Clean
            Application.EnableEvents = False
            BTS_ComponentChoices
        End If
    ElseIf Sh.Name = BTS_T(1) Then
        If Not Intersect(Target, Sh.Range("A2:F2")) Is Nothing Then
            On Error GoTo Clean
            Application.EnableEvents = False
            Dim chosen As String, autoText As String
            chosen = CodeFrom(Sh)
            If chosen <> "" And IsAssemblyCode(chosen) Then
                autoText = AutomaticName(chosen)
                If autoText <> "" Then Sh.Range("B4").Value = autoText
            End If
        End If
        If Not Intersect(Target, Sh.Range("A9:A48")) Is Nothing Then
            On Error GoTo Clean
            Application.EnableEvents = False
            Dim r As Long
            For r = 9 To 48
                If CStr(Sh.Cells(r, 1).Value) <> "" Then Sh.Cells(r, 2).Value = ItemName(CStr(Sh.Cells(r, 1).Value))
            Next r
        End If
    End If
Clean:
    Application.EnableEvents = True
End Sub