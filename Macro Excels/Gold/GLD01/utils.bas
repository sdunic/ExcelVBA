Attribute VB_Name = "utils"
'pass je gold1950

Function getLastRow(column As String) As Long
    Dim sht As Worksheet
    Set sht = ActiveSheet
    getLastRow = sht.Cells(sht.rows.Count, column).End(xlUp).row + 1
End Function

Function getUserName() As String
    'domenski korisnik
    getUserName = Environ$("username")

    'ime i prezime
    'GetUserName = Application.UserName
End Function

Function getDateString(val As Date) As String
    If val = 0 Then
        getDateString = "NULL"
    Else
        getDateString = "to_date(''" & Format(Day(val), "00") & "-" & Format(Month(val), "00") & "-" & Year(val) & "'',''DD-MM-YYYY'')"
    End If
End Function

Function getPriceValue(val As Variant) As String
    If IsEmpty(val) Then
        getPriceValue = "NULL"
    Else
        getPriceValue = Replace(CStr(CDbl(Replace(val, ".", ","))), ",", ".")
    End If
End Function

Function getString(val As Variant, val2 As Variant, val3 As Variant) As String
    If IsNull(val) Then
        getString = ""
    Else
        If val2 > 0 And val3 = 0 Then
            getString = CStr(val)
        ElseIf val2 > 0 And val3 = 1 Then
            getString = CStr(Format(val, "#.00"))
        Else
            getString = ""
        End If
    End If
End Function

Function getDateFormat(val As Variant)
    Dim tmpVal1, tmpVal2 As Variant
    
    getDateFormat = val
    
    tmpVal1 = val
    tmpVal2 = Replace(val, ".", "")
    
    On Error Resume Next
    tmpVal1 = CDate(val)
    tmpVal2 = Replace(CDate(val), ".", "")

    If Len(tmpVal1) - Len(tmpVal2) = 3 Then
        getDateFormat = Format(CDate(tmpVal1), "dd.mm.yyyy")
    End If

    getDateFormat = CStr(getDateFormat)

End Function


Sub addComment(cell As Variant, old As Variant, future As Variant)
    'old = stari nabavni uvjet
    'future = nabavni uvjet u buduænosti
    
    Dim info As String
    
    info = ""
    If Len(old) > 0 And globals.getOldCond Then
        info = info & "PRETHODNI " & old & Chr(10)
    End If
    If Len(future) > 0 And globals.getFutureCond Then
        info = info & "BUDUÆI " & future
    End If
    
    If Len(info) > 0 Then
        If Range(cell).Comment Is Nothing Then
            Range(cell).addComment
        Else
            Range(cell).Comment.Delete
            Range(cell).addComment
        End If
        
        Range(cell).Comment.Visible = False
        Range(cell).Comment.text text:=info
    Else
        If Range(cell).Comment Is Nothing Then
            'do nothing
        Else
            Range(cell).Comment.Delete
        End If
    End If
End Sub

Function validatePurchaseConditions() As Boolean

    validatePurchaseConditions = True
    Dim msg As String
    Dim i As Long
    For i = 6 To getLastRow(cfg.getcTNUVAL601)
        Application.Goto Range(CStr(cfg.getcTNUVAL601) & CStr(i) & ":" & CStr(cfg.getcTNUDFIN601) & CStr(i)), True
        Range(CStr(cfg.getcTNUVAL601) & CStr(i) & ":" & CStr(cfg.getcTNUDFIN601) & CStr(i)).Select
        msg = validatePurchaseCondition(Range(cfg.getcTNUVAL601 & i), Range(cfg.getcTNUUAPP601 & i), Range(cfg.getcTNUDDEB601 & i), Range(cfg.getcTNUDFIN601 & i), "601")
        If msg <> "OK" Then
            MsgBox msg, vbCritical, "Geška"
            validatePurchaseConditions = False
            Exit For
        End If
        
        Application.Goto Range(CStr(cfg.getcTNUVAL602) & CStr(i) & ":" & CStr(cfg.getcTNUDFIN602) & CStr(i)), True
        Range(CStr(cfg.getcTNUVAL602) & CStr(i) & ":" & CStr(cfg.getcTNUDFIN602) & CStr(i)).Select
        msg = validatePurchaseCondition(Range(cfg.getcTNUVAL602 & i), Range(cfg.getcTNUUAPP602 & i), Range(cfg.getcTNUDDEB602 & i), Range(cfg.getcTNUDFIN602 & i), "602")
        If msg <> "OK" Then
            MsgBox msg, vbCritical, "Geška"
            validatePurchaseConditions = False
            Exit For
        End If
        
        Application.Goto Range(CStr(cfg.getcTNUVAL603) & CStr(i) & ":" & CStr(cfg.getcTNUDFIN603) & CStr(i)), True
        Range(CStr(cfg.getcTNUVAL603) & CStr(i) & ":" & CStr(cfg.getcTNUDFIN603) & CStr(i)).Select
        msg = validatePurchaseCondition(Range(cfg.getcTNUVAL603 & i), Range(cfg.getcTNUUAPP603 & i), Range(cfg.getcTNUDDEB603 & i), Range(cfg.getcTNUDFIN603 & i), "603")
        If msg <> "OK" Then
            MsgBox msg, vbCritical, "Geška"
            validatePurchaseConditions = False
            Exit For
        End If
        
        Application.Goto Range(CStr(cfg.getcTNUVAL604) & CStr(i) & ":" & CStr(cfg.getcTNUDFIN604) & CStr(i)), True
        Range(CStr(cfg.getcTNUVAL604) & CStr(i) & ":" & CStr(cfg.getcTNUDFIN604) & CStr(i)).Select
        msg = validatePurchaseCondition(Range(cfg.getcTNUVAL604 & i), Range(cfg.getcTNUUAPP604 & i), Range(cfg.getcTNUDDEB604 & i), Range(cfg.getcTNUDFIN604 & i), "604")
        If msg <> "OK" Then
            MsgBox msg, vbCritical, "Geška"
            validatePurchaseConditions = False
            Exit For
        End If
        
        Application.Goto Range(CStr(cfg.getcTNUVAL605) & CStr(i) & ":" & CStr(cfg.getcTNUDFIN605) & CStr(i)), True
        Range(CStr(cfg.getcTNUVAL605) & CStr(i) & ":" & CStr(cfg.getcTNUDFIN605) & CStr(i)).Select
        msg = validatePurchaseCondition(Range(cfg.getcTNUVAL605 & i), Range(cfg.getcTNUUAPP605 & i), Range(cfg.getcTNUDDEB605 & i), Range(cfg.getcTNUDFIN605 & i), "605")
        If msg <> "OK" Then
            MsgBox msg, vbCritical, "Geška"
            validatePurchaseConditions = False
            Exit For
        End If
        
        Application.Goto Range(CStr(cfg.getcTNUVAL606) & CStr(i) & ":" & CStr(cfg.getcTNUDFIN606) & CStr(i)), True
        Range(CStr(cfg.getcTNUVAL606) & CStr(i) & ":" & CStr(cfg.getcTNUDFIN606) & CStr(i)).Select
        msg = validatePurchaseCondition(Range(cfg.getcTNUVAL606 & i), Range(cfg.getcTNUUAPP606 & i), Range(cfg.getcTNUDDEB606 & i), Range(cfg.getcTNUDFIN606 & i), "606")
        If msg <> "OK" Then
            MsgBox msg, vbCritical, "Geška"
            validatePurchaseConditions = False
            Exit For
        End If
    Next

End Function

Function validateNCValues(nc As Range, app As Range, nnc As Range, exnnc As Range, ddeb As Range, dfin As Range, tcp) As Boolean
    validateNCValues = "OK"
End Function

Function validatePurchaseCondition(val As Range, app As Range, ddeb As Range, dfin As Range, disc As String) As String
    
    validatePurchaseCondition = "OK"
    If val.Font.Color = 255 Or app.Font.Color = 255 Or ddeb.Font.Color = 255 Or dfin.Font.Color = 255 Then
        If Len(val.Value) > 0 And Not Application.WorksheetFunction.IsNumber(val.Value) Then
            validatePurchaseCondition = "Potrebno je upisati ispravan podatak za vrijednost popusta - brojèani podatak! [ERRCODE: " & disc & "VAL]"
            Exit Function
        End If
        If app.Value <> "%" And app.Value <> "iznos" Then
            validatePurchaseCondition = "Potrebno je upisati ispravan podatak za jedinicu popusta - ili % ili iznos! [ERRCODE: " & disc & "APP]"
            Exit Function
        End If
        If Len(ddeb.Value) = 0 Or Not IsDate(ddeb.Value) Then
            validatePurchaseCondition = "Potrebno je upisati ispravan oblik datuma! [ERRCODE: " & disc & "DDEB]"
            Exit Function
        End If
         If Len(dfin.Value) = 0 Or Not IsDate(dfin.Value) Then
            validatePurchaseCondition = "Potrebno je upisati ispravan oblik datuma! [ERRCODE: " & disc & "DFIN]"
            Exit Function
        End If
    End If
    
End Function

Sub setTNUNNC(tnunnc As String, oldTnunnc As String, row As Long)

    If tnunnc = "NETO" Or tnunnc = "NAC" Then
        ActiveWorkbook.Sheets(2).Select
        setDiscount row
    ElseIf oldTnunnc = "NETO" Or oldTnunnc = "NAC" Then
        ActiveWorkbook.Sheets(4).Select
        clearDiscount row
    End If

End Sub

Sub setDiscount(row As Long)

    Dim paDdeb, paDfin As String
    paDdeb = cfg.getcTNUPADDEB
    paDfin = cfg.getcTNUPADFIN

    setNetoZeroDiscount cfg.getcTNUVAL601, cfg.getcTNUDDEB601, cfg.getcTNUDFIN601, CStr(paDdeb), CStr(paDfin), row
    setNetoZeroDiscount cfg.getcTNUVAL602, cfg.getcTNUDDEB602, cfg.getcTNUDFIN602, CStr(paDdeb), CStr(paDfin), row
    setNetoZeroDiscount cfg.getcTNUVAL603, cfg.getcTNUDDEB603, cfg.getcTNUDFIN603, CStr(paDdeb), CStr(paDfin), row
    setNetoZeroDiscount cfg.getcTNUVAL604, cfg.getcTNUDDEB604, cfg.getcTNUDFIN604, CStr(paDdeb), CStr(paDfin), row
    setNetoZeroDiscount cfg.getcTNUVAL605, cfg.getcTNUDDEB605, cfg.getcTNUDFIN605, CStr(paDdeb), CStr(paDfin), row
    setNetoZeroDiscount cfg.getcTNUVAL606, cfg.getcTNUDDEB606, cfg.getcTNUDFIN606, CStr(paDdeb), CStr(paDfin), row

End Sub

Sub setNetoZeroDiscount(colVal As String, colDdeb As String, colDfin As String, colPaDdeb As String, colPaDfin As String, row As Long)
    If Len(Range(colVal & row).Value) > 0 Then
        Range(colVal & row).Value = 0
        Range(colDdeb & row).Value = Range(colPaDdeb & row).Value
        Range(colDfin & row).Value = Range(colPaDfin & row).Value
    End If
End Sub

Sub clearDiscount(row As Long)

    ActiveWorkbook.Sheets(4).Range("A:A").Select
    
    clearNetoZeroDiscount cfg.getcTNUVAL601, cfg.getcTNUDDEB601, cfg.getcTNUDFIN601, row
    clearNetoZeroDiscount cfg.getcTNUVAL602, cfg.getcTNUDDEB602, cfg.getcTNUDFIN602, row
    clearNetoZeroDiscount cfg.getcTNUVAL603, cfg.getcTNUDDEB603, cfg.getcTNUDFIN603, row
    clearNetoZeroDiscount cfg.getcTNUVAL604, cfg.getcTNUDDEB604, cfg.getcTNUDFIN604, row
    clearNetoZeroDiscount cfg.getcTNUVAL605, cfg.getcTNUDDEB605, cfg.getcTNUDFIN605, row
    clearNetoZeroDiscount cfg.getcTNUVAL606, cfg.getcTNUDDEB606, cfg.getcTNUDFIN606, row

End Sub

Sub clearNetoZeroDiscount(colVal As String, colDdeb As String, colDfin As String, row As Long)
    If Len(ActiveWorkbook.Sheets(2).Range(colVal & row).Value) > 0 Then
        Set cell = Selection.Find(What:=Cells(row, ActiveWorkbook.Sheets(2).Range(colVal & row).column).address, After:=ActiveCell, LookIn:=xlFormulas, _
            LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
            MatchCase:=False, SearchFormat:=False)
            
        ActiveWorkbook.Sheets(2).Range(colVal & row).Value = CDbl(ActiveWorkbook.Sheets(4).Range("C" & cell.row).Value)
        ActiveWorkbook.Sheets(2).Range(colDdeb & row).Value = ActiveWorkbook.Sheets(4).Range("C" & cell.row + 2).Value
        ActiveWorkbook.Sheets(2).Range(colDfin & row).Value = ActiveWorkbook.Sheets(4).Range("C" & cell.row + 3).Value
        
    End If
End Sub


Sub allowEventHandling()

    globals.setAllowEventHandling (True)
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault

End Sub

