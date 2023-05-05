Attribute VB_Name = "functions"
Sub insertLog(operation As String, parameters As String, sqlquery As String)
    Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        SQLStr = queries.getLog(db.getDocType, db.getDocName, db.getDocVersion, utils.getUserName, operation, parameters, Replace(sqlquery, "'", """"))
        'Debug.Print SQLstr
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLStr, Cn, adOpenStatic
        
        Cn.Close
        Set Cn = Nothing
End Sub

Sub printRows()
    utils.docUnlock
    utils.sendToPrinter ("$B$3:$L" & "$" & utils.getLastRow("B") - 1)
    insertLog "print_doc", "{ docType: " & Range("B3").Value & " }", ""
    utils.docLock
End Sub

Sub clearDoc()
    Dim sht As Worksheet
    Set sht = ActiveSheet
    Range("L2").Value = utils.getUserName
    Range("C5:L5").ClearContents
    Range("B8:F8").ClearContents
    Range("B11:L13").ClearContents
    Rows("16:" & utils.getLastRow("B") + 20).Select
    Selection.Delete Shift:=xlUp
    Range("B5").Activate
End Sub

Sub initDocument()
    utils.docUnlock
    ans = MsgBox("Jeste li sigurni da želite poèistiti dokument?", vbYesNo, "Upozorenje")
    If ans = 6 Then
        insertLog "clear_doc", "", ""
        clearDoc
        Range("B5").ClearContents
    ElseIf ans = 7 Then
        'NO
    End If
    utils.docLock
End Sub

Sub getData()

    utils.docUnlock
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    If Len(Trim(Range("B5").Value)) > 0 Then
        clearDoc
                       
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        SQLStr = queries.getHeader(Trim(Range("B5").Value))
        insertLog "load_doc_header", _
        "{ orderId: " & Range("B5").Value _
        & " }", CStr(SQLStr)
        'Debug.Print SQLStr
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLStr, Cn, adOpenStatic
        
        Dim discount As Double
        
        Do Until rs.EOF = True
            Range("C5").Value = rs(1) 'datum i vrijeme narudžbe
            Range("D5").Value = rs(2) 'naruèeno od
            Range("E5").Value = rs(4) 'oznaka ugovora
            Range("F5").Value = rs(3) 'šifra kupca
            Range("G5").Value = rs(7) 'naziv kupca
            Range("I5").Value = Application.WorksheetFunction.Proper(rs(9) + ", " + rs(10)) 'adresa isporuke
            
            Range("B8").Value = rs(11) 'valuta
            Range("C8").Value = rs(12) 'konsignacija
            Range("D8").Value = rs(13) 'datum isporuke
            Range("E8").Value = rs(14) 'ruta
            Range("F8").Value = rs(15) 'status
            
            Range("B11").Value = rs(16) 'komentar
            
            rs.MoveNext
        Loop
        
    
    
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        SQLStr = queries.getDetails(Trim(Range("B5").Value))
        'Debug.Print SQLStr
        
        insertLog "load_doc_details", _
        "{ orderId: " & Range("B5").Value _
        & " }", CStr(SQLStr)
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLStr, Cn, adOpenStatic
        
        Dim ukupnaKolicina As Double
        
        ukupnaKolicina = 0
       
        row = 16
        Do Until rs.EOF = True
            Range("B" & row).Value = rs(0) 'šifra artikla
            Range("C" & row).Value = rs(1) 'naziv artikla
            Range("D" & row).Value = rs(2) 'lv
            Range("E" & row).Value = LCase(rs(3)) 'jedinica
            Range("F" & row).Value = rs(4) 'stopa PDV-a
            Range("G" & row).Value = rs(5) 'kolièina
            Range("H" & row).Value = rs(6) 'koeficijent
            Range("I" & row).Value = rs(7) 'kolièina njz
            Range("J" & row).Value = rs(8) 'cijena
            Range("J" & row).NumberFormat = "#,##0.00 €"
            Range("K" & row).Value = LCase(rs(9)) 'jedinica aplik.
            Range("L" & row).Value = rs(10) 'iznos
            Range("L" & row).NumberFormat = "#,##0.00 €"
            
            ukupnaKolicina = ukupnaKolicina + rs(7)
           
            row = row + 1
            rs.MoveNext
        Loop
        
        
        createFooter
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        SQLStr = queries.getFooter(Trim(Range("B5").Value))
        'Debug.Print SQLStr
        
        insertLog "load_doc_footer", _
        "{ orderId: " & Range("B5").Value _
        & " }", CStr(SQLStr)
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLStr, Cn, adOpenStatic
        
        
        Dim sveukupno, sveukupnoPDV As Double
        sveukupno = 0
        svsveukupnoPDV = 0
        
        row = utils.getLastRow("L")
        Do Until rs.EOF = True
            Range("B" & row).Value = rs(0) 'stopa pdv-a
            Range("C" & row).Value = rs(1) 'osnovica
            Range("D" & row).Value = rs(2) 'iznos s pdv-om
            
            sveukupno = sveukupno + rs(1)
            svsveukupnoPDV = svsveukupnoPDV + rs(1) + rs(2)
            
            row = row + 1
            rs.MoveNext
        Loop
                
        Range("J" & utils.getLastRow("L")).Value = ukupnaKolicina
        Range("K" & utils.getLastRow("L")).Value = sveukupno
        Range("L" & utils.getLastRow("L")).Value = svsveukupnoPDV
        
        clearBorders
                       
        rs.Close
        Set rs = Nothing
        Cn.Close
        Set Cn = Nothing
    Else
        MsgBox "Potrebno je upisati broj narudžbe!", vbOKOnly, "Informacija"
        Range("B5").Select
    End If
    
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    utils.docLock

End Sub


Sub createFooter()

    Dim redak As Long
    redak = utils.getLastRow("L")

    Range("A" & redak & ":M" & redak + 2).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    Range("B" & redak + 3).Value = "Stopa PDV-a"
    Range("C" & redak + 3).Value = "Osnovica"
    Range("D" & redak + 3).Value = "Iznos PDV-a"
    Range("B" & redak + 3 & ":D" & redak + 3).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 4464858
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    With Selection.Font
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    
    Range("J" & redak + 3).Value = "Ukupna  kolièina"
    Range("K" & redak + 3).Value = "Iznos"
    Range("K" & redak + 3).Value = "Sveukupno"
    Range("L" & redak + 3).Value = "Sveukupno s PDV-om"
    Range("D" & redak + 3).Select
    Selection.Copy
    Range("J" & redak + 3 & ":L" & redak + 3).Select
    Selection.PasteSpecial Paste:=xlPasteFormats, operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("B" & redak + 4 & ":B" & redak + 7).Select
    Selection.NumberFormat = "#,##0.00"
    Range("C" & redak + 4 & ":D" & redak + 7).Select
    Selection.NumberFormat = "#,##0.00 €"
    
    Range("J" & redak + 4).Select
    Selection.NumberFormat = "#,##0.00"
    Range("K" & redak + 4 & ":L" & redak + 4).Select
    Selection.NumberFormat = "#,##0.00 €"
    
    Range("C" & redak + 4 & ":D" & redak + 7).Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("J" & redak + 4 & ":L" & redak + 4).Select
    With Selection
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
End Sub

Sub clearBorders()
    
    row = utils.getLastRow("E")
    rowRight = utils.getLastRow("L") - 1
    rowLeft = utils.getLastRow("B") - 1
    
    Range("E" & row & ":I" & rowRight).Select
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    Range("E" & rowRight + 1 & ":L" & rowLeft).Select
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone

End Sub


