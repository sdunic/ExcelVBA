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
        
        'TNC
        utils.setPrice row, cfg.getColTNC_Datum, rs(cfg.getRsTNC_Datum), rs(cfg.getRsTNC_DatumKraja), cfg.getColTNC_Cijena, rs(cfg.getRsTNC_Cijena), cfg.getColTNC_NovaCijena, cfg.getColTNC_Indeks, rs(cfg.getRsTNC_Ntar), cfg.getColBrojPromjena
        'TNC A
        utils.setPrice row, cfg.getColTNC_ADatum, rs(cfg.getRsTNC_ADatum), rs(cfg.getRsTNC_ADatumKraja), cfg.getColTNC_ACijena, rs(cfg.getRsTNC_ACijena), cfg.getColTNC_ANovaCijena, cfg.getColTNC_AIndeks, rs(cfg.getRsTNC_ANtar), cfg.getColBrojPromjena
        'TNC B
        utils.setPrice row, cfg.getColTNC_BDatum, rs(cfg.getRsTNC_BDatum), rs(cfg.getRsTNC_BDatumKraja), cfg.getColTNC_BCijena, rs(cfg.getRsTNC_BCijena), cfg.getColTNC_BNovaCijena, cfg.getColTNC_BIndeks, rs(cfg.getRsTNC_BNtar), cfg.getColBrojPromjena
        'TNC C
        utils.setPrice row, cfg.getColTNC_CDatum, rs(cfg.getRsTNC_CDatum), rs(cfg.getRsTNC_CDatumKraja), cfg.getColTNC_CCijena, rs(cfg.getRsTNC_CCijena), cfg.getColTNC_CNovaCijena, cfg.getColTNC_CIndeks, rs(cfg.getRsTNC_CNtar), cfg.getColBrojPromjena
        'TNC D
        utils.setPrice row, cfg.getColTNC_DDatum, rs(cfg.getRsTNC_DDatum), rs(cfg.getRsTNC_DDatumKraja), cfg.getColTNC_DCijena, rs(cfg.getRsTNC_DCijena), cfg.getColTNC_DNovaCijena, cfg.getColTNC_DIndeks, rs(cfg.getRsTNC_DNtar), cfg.getColBrojPromjena
        'TNC S
        utils.setPrice row, cfg.getColTNC_SDatum, rs(cfg.getRsTNC_SDatum), rs(cfg.getRsTNC_SDatumKraja), cfg.getColTNC_SCijena, rs(cfg.getRsTNC_SCijena), cfg.getColTNC_SNovaCijena, cfg.getColTNC_SIndeks, rs(cfg.getRsTNC_SNtar), cfg.getColBrojPromjena
        'TNC KAMP
        utils.setPrice row, cfg.getColTNC_KAMPDatum, rs(cfg.getRsTNC_KAMPDatum), rs(cfg.getRsTNC_KAMPDatumKraja), cfg.getColTNC_KAMPCijena, rs(cfg.getRsTNC_KAMPCijena), cfg.getColTNC_KAMPNovaCijena, cfg.getColTNC_KAMPIndeks, rs(cfg.getRsTNC_KAMPNtar), cfg.getColBrojPromjena
        
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
        utils.setChangedItem i, cfg.getColTNC_Cijena, cfg.getColTNC_NovaCijena, cfg.getColTNC_Indeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColTNC_ACijena, cfg.getColTNC_ANovaCijena, cfg.getColTNC_AIndeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColTNC_BCijena, cfg.getColTNC_BNovaCijena, cfg.getColTNC_BIndeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColTNC_CCijena, cfg.getColTNC_CNovaCijena, cfg.getColTNC_CIndeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColTNC_DCijena, cfg.getColTNC_DNovaCijena, cfg.getColTNC_DIndeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColTNC_SCijena, cfg.getColTNC_SNovaCijena, cfg.getColTNC_SIndeks, cfg.getColBrojPromjena
        utils.setChangedItem i, cfg.getColTNC_KAMPCijena, cfg.getColTNC_KAMPNovaCijena, cfg.getColTNC_KAMPIndeks, cfg.getColBrojPromjena
    Next i
        
    Sheets(3).Select
    Range(cfg.getColSifraArtikla & "3:" & cfg.getColBrojPromjena & utils.getLastRow(cfg.getColSifraArtikla)).ClearContents
    
    Sheets(2).Select
    
    tmpLastRow = utils.getLastRow(cfg.getColSifraArtikla)
    ActiveSheet.Range("$" & cfg.getColSifraArtikla & "$4:$" & cfg.getColBrojPromjena & "$" & tmpLastRow).AutoFilter Field:=51, Criteria1:=">0"
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
    ActiveSheet.Range("$" & cfg.getColSifraArtikla & "$4:$" & cfg.getColBrojPromjena & "$" & tmpLastRow).AutoFilter Field:=51
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
            
            'TNC - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColTNC_Cijena & "3").Value), Range(cfg.getColTNC_Datum & i).Value, CStr(Range(cfg.getColTNC_Cijena & i).Value), CStr(Range(cfg.getColTNC_NovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))

            'TNC A - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColTNC_ACijena & "3").Value), Range(cfg.getColTNC_ADatum & i).Value, CStr(Range(cfg.getColTNC_ACijena & i).Value), CStr(Range(cfg.getColTNC_ANovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'TNC B - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColTNC_BCijena & "3").Value), Range(cfg.getColTNC_BDatum & i).Value, CStr(Range(cfg.getColTNC_BCijena & i).Value), CStr(Range(cfg.getColTNC_BNovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'TNC C - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColTNC_CCijena & "3").Value), Range(cfg.getColTNC_CDatum & i).Value, CStr(Range(cfg.getColTNC_CCijena & i).Value), CStr(Range(cfg.getColTNC_CNovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'TNC D - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColTNC_DCijena & "3").Value), Range(cfg.getColTNC_DDatum & i).Value, CStr(Range(cfg.getColTNC_DCijena & i).Value), CStr(Range(cfg.getColTNC_DNovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'TNC S - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColTNC_SCijena & "3").Value), Range(cfg.getColTNC_SDatum & i).Value, CStr(Range(cfg.getColTNC_SCijena & i).Value), CStr(Range(cfg.getColTNC_SNovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
            'TNC KAMP - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(CStr(Range(cfg.getColTNC_KAMPCijena & "3").Value), Range(cfg.getColTNC_KAMPDatum & i).Value, CStr(Range(cfg.getColTNC_KAMPCijena & i).Value), CStr(Range(cfg.getColTNC_KAMPNovaCijena & i).Value), Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
            
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

