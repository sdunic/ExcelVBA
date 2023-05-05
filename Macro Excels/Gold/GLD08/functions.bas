Attribute VB_Name = "functions"
Sub insertLog(operation As String, parameters As String, sqlquery As String)
    Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        SQLstr = queries.getLog(db.getDocType, db.getDocName, db.getDocVersion, utils.getUserName, operation, parameters, Replace(sqlquery, "'", """"))
        'Debug.Print SQLstr
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLstr, Cn, adOpenStatic
        
        Cn.Close
        Set Cn = Nothing
End Sub

Sub loadSearch()
    frmSearch.Show
End Sub

Sub loadPrixes()

    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    globals.setAllowEventHandling False
    
    If (Len(range("C8").Value) > 0 Or Len(range("C9").Value) > 0 Or Len(range("C10").Value) > 0) Then

        db.getConnectionString
        
        If Len(range("C8").Value) > 0 Then
            ntar = Split(range("C8").Value, " - ")(0)
        End If
        
        If Len(range("C10").Value) > 0 Then
            arvcexr = Split(range("C10").Value, " - ")(0)
        End If
        
        If Len(range("C12").Value) > 0 Then
            msnode = Split(range("C12").Value, " - ")(0)
        End If
        
        
        ActiveWorkbook.Sheets(2).Select
        utils.clearValidation ("W5:W" & utils.getLastRow("W"))
        range("B5:W" & utils.getLastRow("B")).Select
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0.499984740745262
        End With
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .ThemeColor = xlThemeColorDark1
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        range("B5:W" & utils.getLastRow("B")).ClearContents
                
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        SQLstr = queries.selectPrices(CStr(ntar), CStr(arvcexr), CStr(msnode), utils.getDateString(Date))
        'Debug.Print (SQLStr)
        
        insertLog "load_prixes", _
        "{ date: " & Date _
        & ", ms: " & Sheets(1).range("C12").Value _
        & ", ntar: " & Sheets(1).range("C8").Value _
        & ", article: " & Sheets(1).range("C10").Value _
        & " }", CStr(SQLstr)
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLstr, Cn, adOpenStatic
        
        If rs.EOF = False Then
        Dim row As Long
        row = 5
        Do Until rs.EOF = True
            'ARTIKL
            range("B" & row).Value = rs(0) 'gold šifra
            range("C" & row).Value = rs(1) 'cinv
            range("D" & row).Value = rs(2) 'barkod
            range("E" & row).Value = rs(3) 'naziv artikla
            
            'ROBNA GRUPA
            range("F" & row).Value = rs(4) 'nivo 1
            range("G" & row).Value = rs(5) 'naziv 1
            range("H" & row).Value = rs(6) 'nivo 2
            range("I" & row).Value = rs(7) 'naziv 2
            range("J" & row).Value = rs(8) 'nivo 3
            range("K" & row).Value = rs(9) 'naziv 3
            range("L" & row).Value = rs(10) 'nivo 4
            range("M" & row).Value = rs(11) 'naziv 4
            range("N" & row).Value = rs(12) 'nivo 5
            range("O" & row).Value = rs(13) 'naziv 5
            
            'CIJENIK I CIJENA
            range("P" & row).Value = rs(14) 'oznaka cjenika
            range("Q" & row).Value = rs(15) 'naziv cjenika
            range("R" & row).Value = Replace(rs(16), " 00:00:00.0000000", "") 'datum od
            range("S" & row).Value = Replace(rs(17), " 00:00:00.0000000", "") 'datum do
            range("T" & row).Value = rs(18) 'cijena
                        
            'Porezna grupa (CTVA) i CEXV
            range("U" & row).Value = rs(19)
            range("V" & row).Value = rs(20)
            
            'U startu postavljamo NE za gašenje cijene
            range("W" & row).Value = "NE"
            
            If rs(21) = 1 Then
                range("B" & row & ":W" & row).Select
                With Selection.Font
                     .Color = -11489280
                     .TintAndShade = 0
                 End With
                 With Selection.Interior
                     .Pattern = xlSolid
                     .PatternColorIndex = xlAutomatic
                     .ThemeColor = xlThemeColorDark1
                     .TintAndShade = -4.99893185216834E-02
                     .PatternTintAndShade = 0
                 End With
            End If
            
            row = row + 1
            rs.MoveNext
        Loop
        
        utils.setValidation ("W5:W" & utils.getLastRow("W"))
        Application.Goto range("E5"), True
        Else
            MsgBox "Pretraga nije dala rezultat!", vbOKOnly, "Informacija"
            ActiveWorkbook.Sheets(1).Select
        End If
        
        rs.Close
        Set rs = Nothing
        Cn.Close
        Set Cn = Nothing
        
    
    Else
        MsgBox "Potrebno je upisati ulazne parametre!", vbOKOnly, "Informacija"
        range("C8").Select
    End If
    
    
    globals.setAllowEventHandling True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault

End Sub


Sub killPrixes()
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    ans = MsgBox("Jeste li sigurni da želite ugasiti oznaèene cijene?!?!", vbYesNo, "Upozorenje")
    
    If ans = 6 Then
        'YES
        Dim i As Long
        lastRow = utils.getLastRow("B")
        
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        SQLstr = queries.selectFich
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLstr, Cn, adOpenStatic
        
        fich = rs(0)
        
        rs.Close
        Set rs = Nothing
        
        'valuta 191 HRK /978 EUR (future)
        valuta = "978"
        
        globals.setRowCount CLng(9999)
        globals.addRowNumber
        SQLprix = ""
        
        barcodes = ""
        cexr = ""
        cinv = ""
        For i = 5 To lastRow - 1
            If (range("W" & i).Value = "DA") Then
                SQLprix = SQLprix + queries.killPrice(range("P" & i).Value, CDate(range("R" & i).Value), CDate(range("S" & i).Value), CStr(range("T" & i).Value), range("B" & i).Value, range("V" & i).Value, range("U" & i).Value, CStr(fich), CStr(valuta))
            
                If (i = utils.getLastRow("B") - 1) Then
                    barcodes = barcodes & "''" & range("D" & i).Value & "''"
                    cexr = cexr & "''" & range("B" & i).Value & "''"
                    cinv = cinv & "''" & range("C" & i).Value & "''"
                Else
                    barcodes = barcodes & "''" & range("D" & i).Value & "'',"
                    cexr = cexr & "''" & range("B" & i).Value & "'',"
                    cinv = cinv & "''" & range("C" & i).Value & "'',"
                End If
            End If
        Next i
        
        If Len(SQLprix) > 0 Then
            'Debug.Print (SQLprix)
            Set rsPrix = CreateObject("ADODB.Recordset")
            rsPrix.Open SQLprix, Cn, adOpenStatic
        End If
        
        insertLog "kill_prixes", _
        "{ cexr: [" & cexr & "]" _
        & ", cinv: [" & cinv & "]" _
        & ", barcodes: [" & barcodes & "]" _
        & " }", CStr(SQLprix)
        
        Set rsPrix = Nothing
        Cn.Close
        Set Cn = Nothing
        
        MsgBox "Cijene su uspješno pogašene u GOLD-u!", vbOKOnly, "Informacija"
        
    ElseIf ans = 7 Then
        'NO
    End If
    
    ' kasnije možemo pokrenuti sa servera program da obradimo insert cijena i nakon toga bi mogli dohvatiti status ažuriranja cijena
    
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

