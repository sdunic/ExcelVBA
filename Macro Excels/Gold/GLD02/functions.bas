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

Sub checkVersion()

    newVersion = utils.checkNewDocumentVersion
    
    If Len(newVersion) > 0 Then
       MsgBox "Dostupna je nova verzija (v" & newVersion & ") dokumenta. Molim preuzmite novu verziju." & vbCrLf & "Aplikacija æe se zatvoriti nakon ove poruke.", vbOKOnly, "Informacija"
       
       Application.ScreenUpdating = True
       Application.Cursor = xlDefault
       
       ActiveWorkbook.Close saveChanges:=False
    Else
       'continue
    End If

End Sub


Sub loadSearch()
    frmSearch.Show
End Sub

Sub tmp()
    globals.setAllowEventHandling True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

Sub loadPrixes()

    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    globals.setAllowEventHandling False
    
    cfg.Init
    
    If Not IsEmpty(Range("C7").Value) Then
        objcint = Split(Range("C7").Value, " - ")(0)
    End If
    If Not IsEmpty(Range("C9").Value) Then
        arvcexr = Split(Range("C9").Value, " - ")(0)
    End If
    If Not IsEmpty(Range("C11").Value) Then
        cfin = Split(Range("C11").Value, " - ")(1)
    End If
    
    barcodes = "-1"
    If Not IsEmpty(Range("H6:H" & utils.getLastRow("H")).Value) Then
        barcodes = ""
        For i = 6 To utils.getLastRow("H")
            If (Len(Range("H" & i).Value) > 0) Then
                If (i = utils.getLastRow("H") - 1) Then
                    barcodes = barcodes & "''" & Range("H" & i).Value & "''"
                Else
                    barcodes = barcodes & "''" & Range("H" & i).Value & "'',"
                End If
            End If
            
        Next i
    End If
    
    If Len(Range("C13").Value) = 0 Or Not IsDate(Range("C13").Value) Then
        MsgBox "Datum novih cijena je obavezno polje!", vbOKOnly, "Greška"
        Range("C13").Activate
        globals.setAllowEventHandling True
        Application.ScreenUpdating = True
        Application.Cursor = xlDefault
        Exit Sub
    End If
  
    Sheets(3).Select
    Range(cfg.getColSifraArtikla & "5:" & cfg.getColBrojPromjena & utils.getLastRow(cfg.getColSifraArtikla)).ClearContents
    
    ActiveWorkbook.Sheets(2).Select
    Application.Goto Range("E5"), True
    Range(cfg.getColSifraArtikla & "5:" & cfg.getColBrojPromjena & utils.getLastRow(cfg.getColSifraArtikla)).ClearContents

    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    SQLStr = queries.selectPrices(CStr(objcint), CStr(arvcexr), CStr(cfin), CStr(barcodes))
    'Debug.Print (SQLStr)
    
    insertLog "load_prixes", _
    "{ date: " & Date _
    & ", ms: " & Sheets(1).Range("C7").Value _
    & ", article: " & Sheets(1).Range("C9").Value _
    & ", supplier: " & Sheets(1).Range("C11").Value _
    & ", barcodes: [" & barcodes & "]" _
    & ", dateFrom: " & Sheets(1).Range("C13").Value _
    & " }", CStr(SQLStr)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLStr, Cn, adOpenStatic
    
    If rs.EOF = False Then
    Dim row As Long
    row = 5
    Do Until rs.EOF = True
        'ARTIKL
        Range(cfg.getColSifraArtikla & row).Value = rs(cfg.getRsSifraArtikla) 'gold šifra
        Range(cfg.getColBarkodArtikla & row).Value = rs(cfg.getRsBarkodArtikla) 'barkod
        Range(cfg.getColNazivArtikla & row).Value = rs(cfg.getRsNazivArtikla) 'naziv artikla
        
        'BRAND I PRINCIPAL
        Range(cfg.getColBrand & row).Value = rs(cfg.getRsBrand) 'brand
        Range(cfg.getColPrincipal & row).Value = rs(cfg.getRsPrincipal) 'principal
        
        'ROBNA GRUPA
        Range(cfg.getColNivo1 & row).Value = rs(cfg.getRsNivo1) 'nivo 1
        Range(cfg.getColNaziv1 & row).Value = rs(cfg.getRsNaziv1) 'naziv 1
        Range(cfg.getColNivo2 & row).Value = rs(cfg.getRsNivo2) 'nivo 2
        Range(cfg.getColNaziv2 & row).Value = rs(cfg.getRsNaziv2) 'naziv 2
        Range(cfg.getColNivo3 & row).Value = rs(cfg.getRsNivo3) 'nivo 3
        Range(cfg.getColNaziv3 & row).Value = rs(cfg.getRsNaziv3) 'naziv 3
        Range(cfg.getColNivo4 & row).Value = rs(cfg.getRsNivo4) 'nivo 4
        Range(cfg.getColNaziv4 & row).Value = rs(cfg.getRsNaziv4) 'naziv 4
        Range(cfg.getColNivo5 & row).Value = rs(cfg.getRsNivo5) 'nivo 5
        Range(cfg.getColNaziv5 & row).Value = rs(cfg.getRsNaziv5) 'naziv

        'ASORTIMAN, OPIS, SVOJSTVO, TSC i POÈETNA CIJENA
        Range(cfg.getColAsortiman & row).Value = rs(cfg.getRsAsortiman) 'asortiman
        Range(cfg.getColTSC & row).Value = utils.getPriceValue(rs(cfg.getRsTSC)) 'tsc
        Range(cfg.getColOpis & row).Value = rs(cfg.getRsOpis) 'opis
        Range(cfg.getColSvojstva & row).Value = rs(cfg.getRsSvojstva) 'svojstva
        Range(cfg.getColPocetnaCijena & row).Value = utils.getPriceValue(rs(cfg.getRsMPC_ACijena)) 'poèetna cijena MPC A
        
        'MPC A
        utils.setPrice row, cfg.getColMPC_ADatum, rs(cfg.getRsMPC_ADatum), cfg.getColMPC_ACijena, rs(cfg.getRsMPC_ACijena), cfg.getColMPC_ANovaCijena, cfg.getColMPC_AIndeks, rs(cfg.getRsMPC_ANtar), cfg.getColBrojPromjena
        'MPC B
        utils.setPrice row, cfg.getColMPC_BDatum, rs(cfg.getRsMPC_BDatum), cfg.getColMPC_BCijena, rs(cfg.getRsMPC_BCijena), cfg.getColMPC_BNovaCijena, cfg.getColMPC_BIndeks, rs(cfg.getRsMPC_BNtar), cfg.getColBrojPromjena
        'MPC C
        utils.setPrice row, cfg.getColMPC_CDatum, rs(cfg.getRsMPC_CDatum), cfg.getColMPC_CCijena, rs(cfg.getRsMPC_CCijena), cfg.getColMPC_CNovaCijena, cfg.getColMPC_CIndeks, rs(cfg.getRsMPC_CNtar), cfg.getColBrojPromjena
        'MPC D
        utils.setPrice row, cfg.getColMPC_DDatum, rs(cfg.getRsMPC_DDatum), cfg.getColMPC_DCijena, rs(cfg.getRsMPC_DCijena), cfg.getColMPC_DNovaCijena, cfg.getColMPC_DIndeks, rs(cfg.getRsMPC_DNtar), cfg.getColBrojPromjena
        'MPC S1
        utils.setPrice row, cfg.getColMPC_S1Datum, rs(cfg.getRsMPC_S1Datum), cfg.getColMPC_S1Cijena, rs(cfg.getRsMPC_S1Cijena), cfg.getColMPC_S1NovaCijena, cfg.getColMPC_S1Indeks, rs(cfg.getRsMPC_S1Ntar), cfg.getColBrojPromjena
        'MPC S2
        utils.setPrice row, cfg.getColMPC_S2Datum, rs(cfg.getRsMPC_S2Datum), cfg.getColMPC_S2Cijena, rs(cfg.getRsMPC_S2Cijena), cfg.getColMPC_S2NovaCijena, cfg.getColMPC_S2Indeks, rs(cfg.getRsMPC_S2Ntar), cfg.getColBrojPromjena
        'MPC S3
        utils.setPrice row, cfg.getColMPC_S3Datum, rs(cfg.getRsMPC_S3Datum), cfg.getColMPC_S3Cijena, rs(cfg.getRsMPC_S3Cijena), cfg.getColMPC_S3NovaCijena, cfg.getColMPC_S3Indeks, rs(cfg.getRsMPC_S3Ntar), cfg.getColBrojPromjena
        'MPC KAMP
        utils.setPrice row, cfg.getColMPC_KAMPDatum, rs(cfg.getRsMPC_KAMPDatum), cfg.getColMPC_KAMPCijena, rs(cfg.getRsMPC_KAMPCijena), cfg.getColMPC_KAMPNovaCijena, cfg.getColMPC_KAMPIndeks, rs(cfg.getRsMPC_KAMPNtar), cfg.getColBrojPromjena
        
        
        
        'Porezna grupa (CTVA) i CEXV
        Range(cfg.getColPoreznaGrupa & row).Value = rs(cfg.getRsPoreznaGrupa)
        Range(cfg.getColCEXV & row).Value = rs(cfg.getRsCEXV)
        
        'Sort poredak
        Range(cfg.getColRedak & row).Value = row - 4 'broj poretka
        
        row = row + 1
        rs.MoveNext
    Loop
    Else
        MsgBox "Pretraga nije dala rezultat!", vbOKOnly, "Informacija"
        ActiveWorkbook.Sheets(1).Select
    End If
    
    rs.Close
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
    
    
    globals.setAllowEventHandling True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault

End Sub

Sub loadChanges()
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    cfg.Init
    
    Dim i As Long
    lastRow = utils.getLastRow(cfg.getColSifraArtikla)
    
    Dim cellCol As Integer
    Dim cellRow As Integer
    
    cellCol = ActiveCell.column
    cellRow = ActiveCell.row
    
    For i = 5 To lastRow - 1
        Range(cfg.getColBrojPromjena & i).ClearContents
        utils.setChangedItem i, cfg.getColMPC_ACijena, cfg.getColMPC_ANovaCijena, cfg.getColMPC_AIndeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColMPC_BCijena, cfg.getColMPC_BNovaCijena, cfg.getColMPC_BIndeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColMPC_CCijena, cfg.getColMPC_CNovaCijena, cfg.getColMPC_CIndeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColMPC_DCijena, cfg.getColMPC_DNovaCijena, cfg.getColMPC_DIndeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColMPC_S1Cijena, cfg.getColMPC_S1NovaCijena, cfg.getColMPC_S1Indeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColMPC_S2Cijena, cfg.getColMPC_S2NovaCijena, cfg.getColMPC_S2Indeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColMPC_S3Cijena, cfg.getColMPC_S3NovaCijena, cfg.getColMPC_S3Indeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColMPC_KAMPCijena, cfg.getColMPC_KAMPNovaCijena, cfg.getColMPC_KAMPIndeks, cfg.getColBrojPromjena
    Next i
        
    Sheets(3).Select
    Range(cfg.getColSifraArtikla & "3:" & cfg.getColBrojPromjena & utils.getLastRow(cfg.getColSifraArtikla)).ClearContents
    
    Sheets(2).Select
    
    tmpLastRow = utils.getLastRow(cfg.getColSifraArtikla)
    ActiveSheet.Range("$" & cfg.getColSifraArtikla & "$4:$" & cfg.getColBrojPromjena & "$" & tmpLastRow).AutoFilter Field:=56, Criteria1:=">0"
    Range(cfg.getColSifraArtikla & "3:" & cfg.getColBrojPromjena & lastRow).Select
    Selection.Copy
    
    Sheets(3).Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    barcodes = ""
    cexr = ""
    If Not IsEmpty(Range(cfg.getColSifraArtikla & "5:" & cfg.getColSifraArtikla & utils.getLastRow(cfg.getColSifraArtikla)).Value) Then
        For i = 5 To utils.getLastRow(cfg.getColSifraArtikla)
            If (Len(Range(cfg.getColSifraArtikla & i).Value) > 0) Then
                If (i = utils.getLastRow(cfg.getColSifraArtikla) - 1) Then
                    barcodes = barcodes & "''" & Range(cfg.getColBarkodArtikla & i).Value & "''"
                    cexr = cexr & "''" & Range(cfg.getColSifraArtikla & i).Value & "''"
                Else
                    barcodes = barcodes & "''" & Range(cfg.getColBarkodArtikla & i).Value & "'',"
                    cexr = cexr & "''" & Range(cfg.getColSifraArtikla & i).Value & "''" & "'',"
                End If
            End If
        Next i
    End If
                       
    insertLog "load_prix_changes", _
    "{ cexr: [" & cexr & "]" _
    & ", barcodes: [" & barcodes & "]" _
    & " }", ""
    
    Sheets(2).Select
    ActiveSheet.Range("$" & cfg.getColSifraArtikla & "$4:$" & cfg.getColBrojPromjena & "$" & tmpLastRow).AutoFilter Field:=56
    Application.Goto Cells(cellRow, cellCol), True
    
    
    Sheets(3).Select
    Range(cfg.getColSifraArtikla & "4:" & cfg.getColBrojPromjena & "4").Select
    Selection.AutoFilter
    Application.Goto Range("Q5"), True

    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

Sub insertChanges()
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    ans = MsgBox("Jeste li sigurni da želite spremiti promjene?", vbYesNo, "Upozorenje")
    
    If ans = 6 Then
        'YES
        cfg.Init
        
        Dim i As Long
        lastRow = utils.getLastRow(cfg.getColSifraArtikla)
        
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        SQLStr = queries.selectFich
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLStr, Cn, adOpenStatic
        
        fich = rs(0)
        
        rs.Close
        Set rs = Nothing
        
        'valuta 191 HRK /978 EUR (future)
        valuta = "978"
        
        globals.setRowCount CLng(9999)
        globals.addRowNumber
        SQLprix = ""
        For i = 5 To lastRow - 1
            
            'MPC A - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColMPC_ACijena & "3").Value), Range(cfg.getColMPC_ADatum & i).Value, CStr(Range(cfg.getColMPC_ACijena & i).Value), CStr(Range(cfg.getColMPC_ANovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'MPC B - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColMPC_BCijena & "3").Value), Range(cfg.getColMPC_BDatum & i).Value, CStr(Range(cfg.getColMPC_BCijena & i).Value), CStr(Range(cfg.getColMPC_BNovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'MPC C - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColMPC_CCijena & "3").Value), Range(cfg.getColMPC_CDatum & i).Value, CStr(Range(cfg.getColMPC_CCijena & i).Value), CStr(Range(cfg.getColMPC_CNovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'MPC D - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColMPC_DCijena & "3").Value), Range(cfg.getColMPC_DDatum & i).Value, CStr(Range(cfg.getColMPC_DCijena & i).Value), CStr(Range(cfg.getColMPC_DNovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'MPC S1 - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColMPC_S1Cijena & "3").Value), Range(cfg.getColMPC_S1Datum & i).Value, CStr(Range(cfg.getColMPC_S1Cijena & i).Value), CStr(Range(cfg.getColMPC_S1NovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'MPC S2 - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColMPC_S2Cijena & "3").Value), Range(cfg.getColMPC_S2Datum & i).Value, CStr(Range(cfg.getColMPC_S2Cijena & i).Value), CStr(Range(cfg.getColMPC_S2NovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'MPC S1 - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColMPC_S3Cijena & "3").Value), Range(cfg.getColMPC_S3Datum & i).Value, CStr(Range(cfg.getColMPC_S3Cijena & i).Value), CStr(Range(cfg.getColMPC_S3NovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'MPC KAMP - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColMPC_KAMPCijena & "3").Value), Range(cfg.getColMPC_KAMPDatum & i).Value, CStr(Range(cfg.getColMPC_KAMPCijena & i).Value), CStr(Range(cfg.getColMPC_KAMPNovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
        
        Next i
        
        'Debug.Print (SQLprix)
        
        Set rsPrix = CreateObject("ADODB.Recordset")
        rsPrix.Open SQLprix, Cn, adOpenStatic
        
        barcodes = ""
        cexr = ""
        If Not IsEmpty(Range(cfg.getColSifraArtikla & "5:" & cfg.getColSifraArtikla & utils.getLastRow(cfg.getColSifraArtikla)).Value) Then
            For i = 5 To utils.getLastRow(cfg.getColSifraArtikla)
                If (Len(Range(cfg.getColSifraArtikla & i).Value) > 0) Then
                    If (i = utils.getLastRow(cfg.getColSifraArtikla) - 1) Then
                        barcodes = barcodes & "''" & Range(cfg.getColBarkodArtikla & i).Value & "''"
                        cexr = cexr & "''" & Range(cfg.getColSifraArtikla & i).Value & "''"
                    Else
                        barcodes = barcodes & "''" & Range(cfg.getColBarkodArtikla & i).Value & "'',"
                        cexr = cexr & "''" & Range(cfg.getColSifraArtikla & i).Value & "''" & "'',"
                    End If
                End If
            Next i
        End If
        
        insertLog "insert_prixes", _
        "{ cexr: [" & cexr & "]" _
        & ", barcodes: [" & barcodes & "]" _
        & " }", CStr(SQLprix)
        
        Set rsPrix = Nothing
        Cn.Close
        Set Cn = Nothing
        
        MsgBox "Cijene su uspješno poslane u GOLD!", vbOKOnly, "Informacija"
        
    ElseIf ans = 7 Then
        'NO
    End If
    
    ' kasnije možemo pokrenuti sa servera program da obradimo insert cijena i nakon toga bi mogli dohvatiti status ažuriranja cijena
    
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

