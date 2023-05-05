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
  
    Sheets(3).Select
    Range(cfg.getColSifraArtikla & "5:" & cfg.getColBrojPromjena & utils.getLastRow(cfg.getColSifraArtikla)).ClearContents
    
    ActiveWorkbook.Sheets(2).Select
    Application.Goto Range("E5"), True
    Range(cfg.getColSifraArtikla & "5:" & cfg.getColBrojPromjena & utils.getLastRow(cfg.getColSifraArtikla)).ClearContents

    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    SQLstr = queries.selectPrices(CStr(objcint), CStr(arvcexr), CStr(cfin))
    'Debug.Print (SQLStr)
    
    insertLog "load_prixes", _
    "{ date: " & Date _
    & ", ms: " & Sheets(1).Range("C7").Value _
    & ", article: " & Sheets(1).Range("C9").Value _
    & ", supplier: " & Sheets(1).Range("C11").Value _
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
        
        'Konzum Hiper
        utils.setPrice row, cfg.getColKonzumHiperDatum, rs(cfg.getRsKonzumHiperDatum), cfg.getColKonzumHiperCijena, rs(cfg.getRsKonzumHiperCijena), cfg.getColKonzumHiperNovaCijena, cfg.getColKonzumHiperIndeks, rs(cfg.getRsKonzumHiperNtar), cfg.getColBrojPromjena
        'Konzum Maxi
        utils.setPrice row, cfg.getColKonzumMaxiDatum, rs(cfg.getRsKonzumMaxiDatum), cfg.getColKonzumMaxiCijena, rs(cfg.getRsKonzumMaxiCijena), cfg.getColKonzumMaxiNovaCijena, cfg.getColKonzumMaxiIndeks, rs(cfg.getRsKonzumMaxiNtar), cfg.getColBrojPromjena
        'Studenac
        utils.setPrice row, cfg.getColStudenacDatum, rs(cfg.getRsStudenacDatum), cfg.getColStudenacCijena, rs(cfg.getRsStudenacCijena), cfg.getColStudenacNovaCijena, cfg.getColStudenacIndeks, rs(cfg.getRsStudenacNtar), cfg.getColBrojPromjena
        
                
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
        utils.setChangedItem i, cfg.getColKonzumHiperCijena, cfg.getColKonzumHiperNovaCijena, cfg.getColKonzumHiperIndeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColKonzumMaxiCijena, cfg.getColKonzumMaxiNovaCijena, cfg.getColKonzumMaxiIndeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColStudenacCijena, cfg.getColStudenacNovaCijena, cfg.getColStudenacIndeks, cfg.getColBrojPromjena
    Next i
        
    Sheets(3).Select
    Range(cfg.getColSifraArtikla & "3:" & cfg.getColBrojPromjena & utils.getLastRow(cfg.getColSifraArtikla)).ClearContents
    
    Sheets(2).Select
    
    tmpLastRow = utils.getLastRow(cfg.getColSifraArtikla)
    ActiveSheet.Range("$" & cfg.getColSifraArtikla & "$4:$" & cfg.getColBrojPromjena & "$" & tmpLastRow).AutoFilter Field:=35, Criteria1:=">0"
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
    ActiveSheet.Range("$" & cfg.getColSifraArtikla & "$4:$" & cfg.getColBrojPromjena & "$" & tmpLastRow).AutoFilter Field:=35
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
            
            'Konzum Hiper - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColKonzumHiperCijena & "3").Value), Range(cfg.getColKonzumHiperDatum & i).Value, CStr(Range(cfg.getColKonzumHiperCijena & i).Value), CStr(Range(cfg.getColKonzumHiperNovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'Konzum Maxi - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColKonzumMaxiCijena & "3").Value), Range(cfg.getColKonzumMaxiDatum & i).Value, CStr(Range(cfg.getColKonzumMaxiCijena & i).Value), CStr(Range(cfg.getColKonzumMaxiNovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'Studenac - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColStudenacCijena & "3").Value), Range(cfg.getColStudenacDatum & i).Value, CStr(Range(cfg.getColStudenacCijena & i).Value), CStr(Range(cfg.getColStudenacNovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
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

