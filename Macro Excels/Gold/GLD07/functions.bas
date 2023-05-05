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

Sub loadSearch()
    frmSearch.Show
End Sub

Sub loadPrixes()

    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    globals.setAllowEventHandling False
    
    If (Len(Range("C8").Value) > 0 Or Len(Range("C9").Value) > 0 Or Len(Range("C10").Value) > 0) And Len(Range("C14").Value) > 0 Then

        db.getConnectionString
        
        If Len(Range("C8").Value) > 0 Then
            ntar = Split(Range("C8").Value, " - ")(0)
        End If
        
        If Len(Range("C9").Value) > 0 Then
            site = Split(Range("C9").Value, " - ")(0)
        End If
        
        If Len(Range("C10").Value) > 0 Then
            arvcexr = Split(Range("C10").Value, " - ")(0)
        End If
        
        If Len(Range("C12").Value) > 0 Then
            msnode = Split(Range("C12").Value, " - ")(0)
        End If
        
        
        ActiveWorkbook.Sheets(2).Select
        Range("B5:V" & utils.getLastRow("B")).Select
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
        Range("B5:V" & utils.getLastRow("B")).ClearContents
                
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        SQLStr = queries.selectPrices(CStr(ntar), CStr(site), CStr(arvcexr), CStr(msnode), utils.getDateString(Sheets(1).Range("C14").Value))
        'Debug.Print (SQLStr)
        
        insertLog "load_prixes", _
        "{ date: " & Date _
        & ", ms: " & Sheets(1).Range("C12").Value _
        & ", ntar: " & Sheets(1).Range("C8").Value _
        & ", site: " & Sheets(1).Range("C19").Value _
        & ", article: " & Sheets(1).Range("C10").Value _
        & ", dateFrom: " & Sheets(1).Range("C14").Value _
        & " }", CStr(SQLStr)
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLStr, Cn, adOpenStatic
        
        If rs.EOF = False Then
        Dim row As Long
        row = 5
        Do Until rs.EOF = True
            'ARTIKL
            Range("B" & row).Value = rs(0) 'gold šifra
            Range("C" & row).Value = rs(1) 'cinv
            Range("D" & row).Value = rs(2) 'barkod
            Range("E" & row).Value = rs(3) 'naziv artikla
            
            'ROBNA GRUPA
            Range("F" & row).Value = rs(4) 'nivo 1
            Range("G" & row).Value = rs(5) 'naziv 1
            Range("H" & row).Value = rs(6) 'nivo 2
            Range("I" & row).Value = rs(7) 'naziv 2
            Range("J" & row).Value = rs(8) 'nivo 3
            Range("K" & row).Value = rs(9) 'naziv 3
            Range("L" & row).Value = rs(10) 'nivo 4
            Range("M" & row).Value = rs(11) 'naziv 4
            Range("N" & row).Value = rs(12) 'nivo 5
            Range("O" & row).Value = rs(13) 'naziv 5
            
            'CIJENIK I CIJENA
            Range("P" & row).Value = rs(14) 'oznaka cjenika
            Range("Q" & row).Value = rs(15) 'naziv cjenika
            Range("R" & row).Value = Replace(rs(16), " 00:00:00.0000000", "") 'datum od
            Range("S" & row).Value = Replace(rs(17), " 00:00:00.0000000", "") 'datum do
            Range("T" & row).Value = rs(18) 'cijena
                        
            'Porezna grupa (CTVA) i CEXV
            Range("U" & row).Value = rs(19)
            Range("V" & row).Value = rs(20)
            
            If rs(21) = 1 Then
                Range("B" & row & ":V" & row).Select
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
        Application.Goto Range("E5"), True
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
        Range("C8").Select
    End If
    
    
    globals.setAllowEventHandling True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault

End Sub

