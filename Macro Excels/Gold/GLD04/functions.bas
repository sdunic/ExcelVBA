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

Sub tmp()
    globals.setAllowEventHandling True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

Sub loadPrixes()
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
    
    SQLstr = queries.selectPrices(CStr(objcint), CStr(arvcexr), CStr(cfin), CStr(barcodes))
    'Debug.Print (SQLStr)
    
    insertLog "load_prixes", _
    "{ date: " & Date _
    & ", ms: " & Sheets(1).Range("C7").Value _
    & ", article: " & Sheets(1).Range("C9").Value _
    & ", supplier: " & Sheets(1).Range("C11").Value _
    & ", barcodes: [" & barcodes & "]" _
    & ", dateFrom: " & Sheets(1).Range("C13").Value _
    & ", dateTo: " & Sheets(1).Range("C15").Value _
    & " }", CStr(SQLstr)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLstr, Cn, adOpenStatic
    
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
        
        'NA
        utils.setPrice row, cfg.getColNA_Datum, rs(cfg.getRsNA_Datum), rs(cfg.getRsNA_DatumKraja), cfg.getColNA_Cijena, rs(cfg.getRsNA_Cijena), cfg.getColNA_NovaCijena, cfg.getColNA_Indeks, rs(cfg.getRsNA_Ntar), cfg.getColBrojPromjena
        'IA
        utils.setPrice row, cfg.getColIA_Datum, rs(cfg.getRsIA_Datum), rs(cfg.getRsIA_DatumKraja), cfg.getColIA_Cijena, rs(cfg.getRsIA_Cijena), cfg.getColIA_NovaCijena, cfg.getColIA_Indeks, rs(cfg.getRsIA_Ntar), cfg.getColBrojPromjena
        'Katalog
        utils.setPrice row, cfg.getColKatalog_Datum, rs(cfg.getRsKatalog_Datum), rs(cfg.getRsKatalog_DatumKraja), cfg.getColKatalog_Cijena, rs(cfg.getRsKatalog_Cijena), cfg.getColKatalog_NovaCijena, cfg.getColKatalog_Indeks, rs(cfg.getRsKatalog_Ntar), cfg.getColBrojPromjena
        'Rasprodaja
        utils.setPrice row, cfg.getColRasprodaja_Datum, rs(cfg.getRsRasprodaja_Datum), rs(cfg.getRsRasprodaja_DatumKraja), cfg.getColRasprodaja_Cijena, rs(cfg.getRsRasprodaja_Cijena), cfg.getColRasprodaja_NovaCijena, cfg.getColRasprodaja_Indeks, rs(cfg.getRsRasprodaja_Ntar), cfg.getColBrojPromjena
        'Istek roka
        utils.setPrice row, cfg.getColIstekRoka_Datum, rs(cfg.getRsIstekRoka_Datum), rs(cfg.getRsIstekRoka_DatumKraja), cfg.getColIstekRoka_Cijena, rs(cfg.getRsIstekRoka_Cijena), cfg.getColIstekRoka_NovaCijena, cfg.getColIstekRoka_Indeks, rs(cfg.getRsIstekRoka_Ntar), cfg.getColBrojPromjena
        
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
End Sub

Sub loadChanges()
    cfg.Init
    
    Dim i As Long
    lastRow = utils.getLastRow(cfg.getColSifraArtikla)
    
    Dim cellCol As Integer
    Dim cellRow As Integer
    
    cellCol = ActiveCell.column
    cellRow = ActiveCell.row
    
    For i = 5 To lastRow - 1
        Range(cfg.getColBrojPromjena & i).ClearContents
        utils.setChangedItem i, cfg.getColNA_Cijena, cfg.getColNA_NovaCijena, cfg.getColNA_Indeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColIA_Cijena, cfg.getColIA_NovaCijena, cfg.getColIA_Indeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColKatalog_Cijena, cfg.getColKatalog_NovaCijena, cfg.getColKatalog_Indeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColRasprodaja_Cijena, cfg.getColRasprodaja_NovaCijena, cfg.getColRasprodaja_Indeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColIstekRoka_Cijena, cfg.getColIstekRoka_NovaCijena, cfg.getColIstekRoka_Indeks, cfg.getColBrojPromjena
    Next i
        
    Sheets(3).Select
    Range(cfg.getColSifraArtikla & "3:" & cfg.getColBrojPromjena & utils.getLastRow(cfg.getColSifraArtikla)).ClearContents
    
    Sheets(2).Select
    
    tmpLastRow = utils.getLastRow(cfg.getColSifraArtikla)
    ActiveSheet.Range("$" & cfg.getColSifraArtikla & "$4:$" & cfg.getColBrojPromjena & "$" & tmpLastRow).AutoFilter Field:=43, Criteria1:=">0"
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
    ActiveSheet.Range("$" & cfg.getColSifraArtikla & "$4:$" & cfg.getColBrojPromjena & "$" & tmpLastRow).AutoFilter Field:=43
    Application.Goto Cells(cellRow, cellCol), True
        
    Sheets(3).Select
    Range(cfg.getColSifraArtikla & "4:" & cfg.getColBrojPromjena & "4").Select
    Selection.AutoFilter
    Application.Goto Range("Q5"), True
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
        For i = 5 To lastRow - 1
            
            'NA - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColNA_Cijena & "3").Value), Range(cfg.getColNA_Datum & i).Value, CStr(Range(cfg.getColNA_Cijena & i).Value), CStr(Range(cfg.getColNA_NovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))

            'IA - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColIA_Cijena & "3").Value), Range(cfg.getColIA_Datum & i).Value, CStr(Range(cfg.getColIA_Cijena & i).Value), CStr(Range(cfg.getColIA_NovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'Katalog - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColKatalog_Cijena & "3").Value), Range(cfg.getColKatalog_Datum & i).Value, CStr(Range(cfg.getColKatalog_Cijena & i).Value), CStr(Range(cfg.getColKatalog_NovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'Rasprodaja - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColRasprodaja_Cijena & "3").Value), Range(cfg.getColRasprodaja_Datum & i).Value, CStr(Range(cfg.getColRasprodaja_Cijena & i).Value), CStr(Range(cfg.getColRasprodaja_NovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'IstekRoka - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColIstekRoka_Cijena & "3").Value), Range(cfg.getColIstekRoka_Datum & i).Value, CStr(Range(cfg.getColIstekRoka_Cijena & i).Value), CStr(Range(cfg.getColIstekRoka_NovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))

            
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

