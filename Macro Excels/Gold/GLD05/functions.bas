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

Sub clearInput()
    Range("C7").ClearContents
    Range("C8").ClearContents
    Range("C10").ClearContents
    Range("C12").ClearContents
    Range("C13").ClearContents
    Range("C15").ClearContents
    Range("C16").ClearContents
End Sub

Sub loadLocalPrixes()

    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    globals.setAllowEventHandling False
    
    If (Len(Range("C10").Value) > 0 Or Len(Range("C12").Value) > 0 Or Len(Range("C13").Value) > 0) And Len(Range("C7").Value) > 0 And Len(Range("C8").Value) > 0 And Len(Range("C8").Value) > 0 And Len(Range("C15").Value) > 0 Then
        
        If (CDate(Range("C15").Value) > Date) Then
        
            cfg.Init
            db.getConnectionString
            
            ntarType = Split(Range("C7").Value, " - ")(0)
            
            If Len(Range("C8").Value) > 0 Then
                sites = Range("C8").Value
            End If
            
            If Len(Range("C10").Value) > 0 Then
                cfin = Split(Range("C10").Value, " - ")(1)
            End If
            
            If Len(Range("C12").Value) > 0 Then
                objcint = Split(Range("C12").Value, " - ")(0)
            End If
            
            If Len(Range("C13").Value) > 0 Then
                arvcexr = Split(Range("C13").Value, " - ")(0)
            End If
            
             If Len(Range("C15").Value) > 0 Then
                datum = Range("C15").Value
            End If
            
            
            Sheets(3).Select
            Range(cfg.getColSifraArtikla & "5:" & cfg.getColCEXV & utils.getLastRow("B")).ClearContents
            
            ActiveWorkbook.Sheets(2).Select
            Application.Goto Range("E5"), True
            Range(cfg.getColSifraArtikla & "5:" & cfg.getColCEXV & utils.getLastRow("B")).ClearContents
        
                    
            Set Cn = CreateObject("ADODB.Connection")
            Cn.ConnectionTimeout = 1000
            Cn.commandtimeout = 1000
            Cn.Open db.getConnectionString
            
            SQLstr = queries.selectLocalPrices(CLng(ntarType), CStr(objcint), CStr(cfin), CStr(arvcexr), CStr(sites), CDate(datum))
            'Debug.Print (SQLStr)
            
            insertLog "load_prixes", _
            "{ date: " & Date _
            & ", objcint: " & CStr(objcint) _
            & ", cfin: " & CStr(cfin) _
            & ", arvcexr: " & CStr(arvcexr) _
            & ", ntarType: " & CStr(ntarType) _
            & ", sites: [" & CStr(sites) & "]" _
            & ", dateFrom: " & Sheets(1).Range("C15").Value _
            & ", dateTo: " & Sheets(1).Range("C16").Value _
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
                
                'BRAND i PRINCIPAL
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
                Range(cfg.getColNaziv5 & row).Value = rs(cfg.getRsNaziv5) 'naziv 5
                
                'OPIS, SVOJSTVO, TSC i poèetna cijena
                Range(cfg.getColTSC & row).Value = utils.getPriceValue(rs(cfg.getRsTSC)) 'tsc
                Range(cfg.getColOpis & row).Value = rs(cfg.getRsOpis) 'opis
                Range(cfg.getColSvojstva & row).Value = rs(cfg.getRsSvojstva) 'svojstva
                
                
                Range(cfg.getColNTAR & row).Value = rs(cfg.getRsNTAR) 'cjenik
                
                If Len(rs(cfg.getRsDatumCijene)) > 0 Then
                    Range(cfg.getColDdeb & row).Value = CDate(rs(cfg.getRsDatumCijene)) 'datum cijene
                End If
                If Len(rs(cfg.getRsDatumKrajaCijene)) > 0 Then
                    Range(cfg.getColDfin & row).Value = CDate(rs(cfg.getRsDatumKrajaCijene)) 'datum kraja cijene
                End If
                Range(cfg.getColPrix & row).Value = rs(cfg.getRsCijena) 'cijena
                
                'Sort redak
                Range(cfg.getColRedak & row).Value = row - 4 'broj poretka
                
                'Porezna grupa (CTVA) i CEXV
                Range(cfg.getColPoreznaGrupa & row).Value = rs(cfg.getRsPoreznaGrupa)
                Range(cfg.getColCEXV & row).Value = rs(cfg.getRsCEXV)
                
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
            
        Else
            MsgBox "Datum od mora biti veæi od današnjeg datuma!", vbOKOnly, "Informacija"
            Range("C15").Select
        End If
        
    Else
        MsgBox "Potrebno je upisati ili dobavljaèa, ili robni èvor, ili artikl te upisati vrstu cjenika, trgovine i datum od kad æe vrijediti cijene!", vbOKOnly, "Informacija"
        Range("C7").Select
    End If
    
    globals.setAllowEventHandling True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault

End Sub

Sub loadChanges()
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    cfg.Init
    
    Dim i As Long
    LastRow = utils.getLastRow(cfg.getColSifraArtikla)
    
    Dim cellCol As Integer
    Dim cellRow As Integer
    
    cellCol = ActiveCell.column
    cellRow = ActiveCell.row
    
    Sheets(3).Select
    Range(cfg.getColSifraArtikla & "3:" & cfg.getColCEXV & utils.getLastRow(cfg.getColSifraArtikla)).ClearContents
    
    Sheets(2).Select
    tmpLastRow = utils.getLastRow(cfg.getColSifraArtikla)
    
    ActiveSheet.Range("$" & cfg.getColSifraArtikla & "$4:$" & cfg.getColCEXV & "$" & tmpLastRow).AutoFilter Field:=23, Criteria1:=">0"
    Range(cfg.getColSifraArtikla & "3:" & cfg.getColCEXV & LastRow).Select
    Selection.Copy
    
    Sheets(3).Select
    Range("B3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    barcodes = ""
    cexr = ""
    ntar = ""
    If Not IsEmpty(Range(cfg.getColSifraArtikla & "5:" & cfg.getColSifraArtikla & utils.getLastRow(cfg.getColSifraArtikla)).Value) Then
        For i = 5 To utils.getLastRow(cfg.getColSifraArtikla)
            If (Len(Range(cfg.getColSifraArtikla & i).Value) > 0) Then
                If (i = utils.getLastRow(cfg.getColSifraArtikla) - 1) Then
                    barcodes = barcodes & "''" & Range(cfg.getColBarkodArtikla & i).Value & "''"
                    cexr = cexr & "''" & Range(cfg.getColSifraArtikla & i).Value & "''"
                    ntar = ntar & "''" & Range(cfg.getColNTAR & i).Value & "''"
                Else
                    barcodes = barcodes & "''" & Range(cfg.getColBarkodArtikla & i).Value & "'',"
                    cexr = cexr & "''" & Range(cfg.getColSifraArtikla & i).Value & "''" & "'',"
                    ntar = ntar & "''" & Range(cfg.getColNTAR & i).Value & "''" & "'',"
                End If
            End If
        Next i
    End If
                       
    insertLog "load_prix_changes", _
    "{ cexr: [" & cexr & "]" _
    & ", barcodes: [" & barcodes & "]" _
    & ", ntar: [" & ntar & "]" _
    & " }", ""
    
    Sheets(2).Select
    ActiveSheet.Range("$" & cfg.getColSifraArtikla & "$4:$" & cfg.getColCEXV & "$" & tmpLastRow).AutoFilter Field:=23
    Application.Goto Cells(cellRow, cellCol), True
    
    
    Sheets(3).Select
    Range(cfg.getColSifraArtikla & "$4:$" & cfg.getColCEXV & "4").Select
    Selection.AutoFilter
    Application.Goto Range("Q5"), True

    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

Sub insertChanges()
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    'potrebno napraviti insert po cjenicima za cijene
    'paziti na datum od i datum do jer svaki cjenik na cijenama ima svoj datum poèetka i kraja
        
    ans = MsgBox("Jeste li sigurni da želite spremiti promjene?", vbYesNo, "Upozorenje")
    
    If ans = 6 Then
        'YES
        cfg.Init
        db.getConnectionString
        
        Dim i As Long
        LastRow = utils.getLastRow(cfg.getColSifraArtikla)
        
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
        For i = 5 To LastRow - 1
            'Cijena - datum, stara cijena, nova cijena - ostalo je sve isto
            SQLprix = SQLprix + queries.getInsertPrix(Range(cfg.getColNTAR & i).Value, _
            CDate(Range(cfg.getColDdeb & i).Value), CDate(Range(cfg.getColDfin & i).Value), CStr(Range(cfg.getColPrix & i).Value), _
            CDate(Sheets(1).Range("C15").Value), CDate(Sheets(1).Range("C16").Value), CStr(Range(cfg.getColNovaCijena & i).Value), _
            Range(cfg.getColSifraArtikla & i).Value, Range(cfg.getColCEXV & i).Value, Range(cfg.getColPoreznaGrupa & i).Value, CStr(fich), CStr(valuta))
        Next i
        
        'Debug.Print
        
        Set rsPrix = CreateObject("ADODB.Recordset")
        rsPrix.Open SQLprix, Cn, adOpenStatic
        
        barcodes = ""
        cexr = ""
        ntar = ""
        If Not IsEmpty(Range(cfg.getColSifraArtikla & "5:" & cfg.getColSifraArtikla & utils.getLastRow(cfg.getColSifraArtikla)).Value) Then
            For i = 5 To utils.getLastRow(cfg.getColSifraArtikla)
                If (Len(Range(cfg.getColSifraArtikla & i).Value) > 0) Then
                    If (i = utils.getLastRow(cfg.getColSifraArtikla) - 1) Then
                        barcodes = barcodes & "''" & Range(cfg.getColBarkodArtikla & i).Value & "''"
                        cexr = cexr & "''" & Range(cfg.getColSifraArtikla & i).Value & "''"
                        ntar = ntar & "''" & Range(cfg.getColNTAR & i).Value & "''"
                    Else
                        barcodes = barcodes & "''" & Range(cfg.getColBarkodArtikla & i).Value & "'',"
                        cexr = cexr & "''" & Range(cfg.getColSifraArtikla & i).Value & "''" & "'',"
                        ntar = ntar & "''" & Range(cfg.getColNTAR & i).Value & "''" & "'',"
                    End If
                End If
            Next i
        End If
        
        insertLog "insert_prixes", _
        "{ cexr: [" & cexr & "]" _
        & ", barcodes: [" & barcodes & "]" _
        & ", ntar: [" & ntar & "]" _
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

