Sub DeleteCheckBoxes()
    Dim rng, cel As Range
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("Get Data").ListObjects("field_list")
    Set rng = tbl.ListColumns("Chooser").DataBodyRange
    
    Dim cb As CheckBox
    For Each cb In ThisWorkbook.Worksheets("Get Data").CheckBoxes
        If Not Intersect(cb.TopLeftCell, rng) Is Nothing Then
            cb.Delete
        End If
    Next
    rng.ClearContents
    
    Dim RName As Name
    For Each RName In Application.ActiveWorkbook.Names
        If InStr(1, RName.Name, "rng_cbx_", vbTextCompare) > 0 Then RName.Delete
    Next
End Sub


Sub UpdateCheckBoxes()
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Worksheets("Get Data").ListObjects("field_list")
    
    Dim rng As Range
    Set rng = tbl.ListColumns("Chooser").DataBodyRange
    rng.NumberFormat = ";;;"
    
    Dim dict, dictx As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Set dictx = CreateObject("Scripting.Dictionary")
    Set dictlv = CreateObject("Scripting.Dictionary")
    
    Dim ref, hpr_1, hpr_2, hpr_3 As Range
    Set ref = tbl.ListColumns("Item Name").DataBodyRange
    Set hpr_1 = tbl.ListColumns("helper_1").DataBodyRange
    Set hpr_2 = tbl.ListColumns("helper_2").DataBodyRange
    Set hpr_3 = tbl.ListColumns("helper_3").DataBodyRange
    Dim hp_1: hp_1 = hpr_1.Value
    Dim hp_2: hp_2 = hpr_2.Value
    Dim hp_3: hp_3 = hpr_3.Value
    
    Dim i As Long
    For i = 1 To rng.Rows.Count
        If i = 1 Then
            dict(i) = 1
            dictx(i) = 1
        Else
            If ref.Cells(i) = "" Then
                dict(i) = dict(i - 1) + 1
                dictx(i) = 1
            Else
                dict(i) = dict(i - 1)
                dictx(i) = dictx(i - 1) + 1
            End If
        End If
        
        If dict(i) = 1 And dictx(i) = 1 Then
            dictlv(i) = 1
        ElseIf dictx(i) = 1 Then
            dictlv(i) = 2
        ElseIf dictx(i) > 1 Then
            dictlv(i) = 3
        End If
        hpr_1(i, 1) = dict(i)
        hpr_2(i, 1) = dictx(i)
        hpr_3(i, 1) = dictlv(i)
    Next
    
    hp_1 = hpr_1
    hp_2 = hpr_2
    hp_3 = hpr_3
    
    Dim hp As Variant
    hp = tbl.ListColumns("helper_1").DataBodyRange
    
    Dim sp, ep As Integer
    For i = 1 To WorksheetFunction.Max(hp)
        sp = Application.Match(i, hp, False)
        ep = Application.Match(i, hp, True)
        ThisWorkbook.Names.Add Name:="rng_cbx_" & Format(i, "00"), RefersTo:=rng.Rows(sp & ":" & ep)
    Next i
    
    Dim tblr, tblrx, cel As Range
    Set tblr = ThisWorkbook.Worksheets("Get Data").ListObjects("field_list").DataBodyRange
    tblr.FormatConditions.Delete
    
    Dim cb As CheckBox
    i = 0
    For Each cel In rng
        i = i + 1
        '--- Add checkboxes ---
        Set cb = ThisWorkbook.Worksheets("Get Data").CheckBoxes.Add(cel.Left + 14.5, cel.Top, 23.25, 16.5)

        With cb
            .Name = "cbx_" & Format(dict(i), "00") & Format(dictx(i), "000")
            .Caption = ""
            .LinkedCell = cel.Address
            If dictlv(i) = 1 Then
                .OnAction = "SelectLv1_Click"
            ElseIf dictlv(i) = 2 Then
                .OnAction = "SelectLv2_Click"
            ElseIf dictlv(i) = 3 Then
                .OnAction = "Mixed_State"
            End If
        End With
        '--- Add conditional formatting ---
        If ref.Rows(i) <> "" Then
            Set tblrx = tblr.Rows(i)
            tblrx.FormatConditions.Add Type:=xlExpression, Formula1:="=$" & Replace(tblrx.Columns(1).Address, "$", "")
            tblrx.FormatConditions(1).Interior.Color = RGB(0, 153, 0)
            tblrx.FormatConditions(1).Font.Color = vbWhite
            tblrx.FormatConditions(1).Font.Bold = True
        End If
    Next
End Sub


Sub SelectLv1_Click()
    Dim ws As Object
    Dim tbl As ListObject
    Dim rng As Range
    Set ws = ThisWorkbook.Worksheets("Get Data")
    Set rng = ws.ListObjects("field_list").ListColumns("Chooser").DataBodyRange
    
    Dim cb As CheckBox
    Dim cbxname As String
    cbxname = Application.Caller
    
    For Each cb In ws.CheckBoxes
        If Not Intersect(cb.TopLeftCell, rng) Is Nothing Then
            If cb.Name <> cbxname Then
                cb.Value = ws.CheckBoxes(cbxname).Value
            End If
        End If
    Next cb
End Sub


Sub SelectLv2_Click()
    Dim ws As Object
    Dim rng As Range
    Set ws = ThisWorkbook.Worksheets("Get Data")
    
    Dim cb As CheckBox
    Dim cbxname As String
    cbxname = Application.Caller
    
    Set rng = ThisWorkbook.Worksheets("Get Data").Range("rng_" & Left(cbxname, 6))
    
    For Each cb In ws.CheckBoxes
        If Not Intersect(cb.TopLeftCell, rng) Is Nothing Then
            If Right(cb.Name, 3) <> "001" Then
                cb.Value = ws.CheckBoxes(cbxname).Value
            End If
        End If
    Next cb
    
    Call Mixed_State
End Sub


Sub Mixed_State()
    Dim ws As Object
    Dim rng As Range
    Set ws = ThisWorkbook.Worksheets("Get Data")
    Set rng = ws.ListObjects("field_list").ListColumns("Chooser").DataBodyRange
    
    '--- Select All checkbox handling ---
    Dim cb As CheckBox
    Set cb = ws.CheckBoxes(Application.Caller)
    
    Const cbxall = "cbx_01001"
    
    If cb.Name <> cbxall And cb.Value <> ws.CheckBoxes(cbxall).Value Then
        If ws.CheckBoxes(cbxall).Value = 1 Then
            'MsgBox "11, cbName = " & cb.Name
            ws.CheckBoxes(cbxall).Value = 0
        ElseIf ws.CheckBoxes(cbxall).Value < 1 And WorksheetFunction.CountIf(rng, True) = rng.Rows.Count - 1 Then
            'MsgBox "12, cbName = " & cb.Name
            ws.CheckBoxes(cbxall).Value = 1
        Else
            'MsgBox "13, cbName = " & cb.Name
            ws.CheckBoxes(cbxall).Value = 0
        End If
    Else
        'MsgBox "14, cbName = " & cb.Name
        ws.CheckBoxes(cbxall).Value = cb.Value
    End If

    '--- Root checkboxes with Select All dependency handling ---
    Dim cbxroot As String
    Dim rngx As Range
    cbxroot = Left(cb.Name, 6) & "001"
    Set rngx = ThisWorkbook.Worksheets("Get Data").Range("rng_" & Left(cb.Name, 6))
    
    If cb.Name <> cbxroot And cb.Value <> ws.CheckBoxes(cbxroot).Value Then
        If ws.CheckBoxes(cbxroot).Value = 1 Then
            'MsgBox "21, cbxroot = " & cbxroot & " cbName = " & cb.Name
            ws.CheckBoxes(cbxroot).Value = 0
        ElseIf ws.CheckBoxes(cbxroot).Value < 1 And WorksheetFunction.CountIf(rngx, True) = rngx.Rows.Count - 1 Then
            If ws.CheckBoxes(cbxall).Value < 1 And WorksheetFunction.CountIf(rng, True) = rng.Rows.Count - 2 Then
                'MsgBox "22, cbxroot = " & cbxroot & " cbName = " & cb.Name
                ws.CheckBoxes(cbxall).Value = 1
                ws.CheckBoxes(cbxroot).Value = 1
            Else
                'MsgBox "23, cbxroot = " & cbxroot & " cbName = " & cb.Name
                ws.CheckBoxes(cbxall).Value = 0
                ws.CheckBoxes(cbxroot).Value = 1
            End If
        Else
            'MsgBox "24, cbxroot = " & cbxroot & " cbName = " & cb.Name
            ws.CheckBoxes(cbxroot).Value = 0
        End If
    Else
        'MsgBox "25, cbxroot = " & cbxroot & " cbName = " & cb.Name
        ws.CheckBoxes(cbxroot).Value = cb.Value
    End If
End Sub
