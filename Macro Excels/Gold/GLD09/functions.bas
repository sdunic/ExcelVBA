Attribute VB_Name = "functions"
Sub insertLog(operation As String, parameters As String, sqlquery As String)
    Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        sqlstr = queries.getLog(db.getDocType, db.getDocName, db.getDocVersion, utils.getUserName, operation, parameters, Replace(sqlquery, "'", """"))
        'Debug.Print SQLstr
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open sqlstr, Cn, adOpenStatic
        
        Cn.Close
        Set Cn = Nothing
End Sub

Sub loadSearchRows()
    cfg.Init
    zaglavlje = cfg.get_zaglavlje
    If IsEmpty(Range(cfg.get_lokacija & zaglavlje)) Or IsEmpty(Range(cfg.get_tipFakture & zaglavlje)) Or IsEmpty(Range(cfg.get_kupac & zaglavlje)) Or IsEmpty(Range(cfg.get_ugovor & zaglavlje)) Or IsEmpty(Range(cfg.get_datumFakture & zaglavlje)) Then
        MsgBox "Potrebno je popuniti zaglavlje fakture!", vbOKOnly, "Upozorenje"
        Range(cfg.get_lokacija & zaglavlje & ":" & cfg.get_datumFakture & zaglavlje).Select
    Else
        If ActiveCell.row >= cfg.get_stavke Then
            frmSearchRows.Show
        Else
            MsgBox "Potrebno je pozicionirait æeliju u stavke fakture!", vbOKOnly, "Upozorenje"
        End If
    End If
End Sub

Sub printRows()
    cfg.Init
    utils.sendToPrinter ("$" & cfg.get_artikl & "$" & cfg.get_stavke - 2 & ":$" & cfg.get_analitickiArtikl & "$" & utils.getLastRow(cfg.get_artikl) - 1)
    insertLog "print_doc", "", ""
End Sub

Sub deleteRow()
    cfg.Init
    currentRow = ActiveCell.row
    If currentRow >= cfg.get_stavke Then
        Application.ScreenUpdating = False
        Rows(currentRow & ":" & currentRow).Select
        Selection.Delete Shift:=xlUp
        Range(cfg.get_artikl & currentRow).Select
        insertLog "delete_row", "", ""
        Application.ScreenUpdating = True
         
    Else
        MsgBox "Potrebno je odabrati stavku koju želite obrisati.", vbOKOnly, "Informacija"
    End If
End Sub

Sub copyRow()
    cfg.Init
    LastRow = utils.getLastRow(cfg.get_artikl)
    currentRow = ActiveCell.row
    If currentRow >= cfg.get_stavke Then
        If Len(Range(cfg.get_artikl & currentRow).Value) > 0 Then
            Application.ScreenUpdating = False
            Range(cfg.get_artikl & currentRow & ":" & cfg.get_analitickiMrezniCvor & currentRow).Copy
            Range(cfg.get_artikl & LastRow & ":" & cfg.get_analitickiMrezniCvor & LastRow).PasteSpecial xlPasteValues
            Application.CutCopyMode = False
            Range(cfg.get_ukupniIznos & LastRow).Select
            insertLog "copy_row", "", ""
            Application.ScreenUpdating = True
        Else
            MsgBox "Stavka mora imati odabran artikl!", vbOKOnly, "Upozorenje"
        End If
    Else
        MsgBox "Potrebno je odabrati stavku koju želite kopiratiu.", vbOKOnly, "Informacija"
    End If
End Sub

Sub loadSearchHeader()
    frmSearchHeader.Show
End Sub

Sub initDocument()
    ans = MsgBox("Jeste li sigurni da želite poèistiti dokument?", vbYesNo, "Upozorenje")
    If ans = 6 Then
        cfg.Init
        Dim sht As Worksheet
        Set sht = ActiveSheet
        Range(cfg.get_korisnik & cfg.get_zaglavlje).Value = utils.getUserName
        Range(cfg.get_lokacija & cfg.get_zaglavlje & ":" & cfg.get_napomena & cfg.get_zaglavlje).ClearContents
        Range(cfg.get_artikl & cfg.get_stavke & ":" & cfg.get_analitickiMrezniCvor & sht.Rows.Count).ClearContents
        Range(cfg.get_lokacija & cfg.get_zaglavlje).Activate
    ElseIf ans = 7 Then
        'NO
    End If
End Sub

Sub insertInvoice()
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    ans = MsgBox("Jeste li sigurni da želite kreirati fakturu u pripremi?", vbYesNo, "Upozorenje")
    
    If ans = 6 Then
        'YES
        cfg.Init
        
        Dim control As Boolean
        control = True
        Dim cexrs() As Variant
                
        Dim i As Long
        LastRow = utils.getLastRow(cfg.get_artikl)
                
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        
        sqlstr = queries.selectMSGID
        Set rsMSGID = CreateObject("ADODB.Recordset")
        rsMSGID.Open sqlstr, Cn, adOpenStatic
        msgid = rsMSGID(0)
        rsMSGID.Close
        Set rsMSGID = Nothing
        
        
        'dohvatiti artikle u listu za kontrolu cexra
        sqlCexrs = queries.getCexrs
        Set rsCexrs = CreateObject("ADODB.Recordset")
        rsCexrs.Open sqlCexrs, Cn, adOpenStatic
        cexrs = rsCexrs.getRows
        rsCexrs.Close
        Set rsCexrs = Nothing
        
        
        fyp = ""
        If ActiveSheet.name = "STORNO GOLD FAKTURE" Then
            fyp = "2"
        End If
        
        
        SQLinsertInvoice = ""
        redakZaglavlja = cfg.get_zaglavlje
        For i = cfg.get_stavke To LastRow - 1
        
     
            control = utils.IsInArray(CStr(Split(Range(cfg.get_artikl & i).Value, " | ")(0)), cexrs)
            
            SQLinsertInvoice = SQLinsertInvoice & queries.insertSIS15(CStr(msgid), (i - cfg.get_stavke) + 1, CStr(Split(Range(cfg.get_lokacija & redakZaglavlja).Value, " | ")(0)), CStr(Range(cfg.get_tipFakture & redakZaglavlja).Value), _
            CStr(Split(Range(cfg.get_kupac & redakZaglavlja).Value, " | ")(0)), CStr(Split(Range(cfg.get_ugovor & redakZaglavlja).Value, " | ")(0)), CStr(Range(cfg.get_datumFakture & redakZaglavlja).Value), _
            CStr(Split(Range(cfg.get_artikl & i).Value, " | ")(0)), CStr(Range(cfg.get_analitickiArtikl & i).Value), CStr(Range(cfg.get_robniCvor & i).Value), _
            CStr(Range(cfg.get_tm & i).Value), CStr(Split(Range(cfg.get_reasonCodeTekst & cfg.get_reasonCodeRedak).Value, " | ")(0)), CStr(Range(cfg.get_korisnik & redakZaglavlja).Value), CStr(Range(cfg.get_napomena & redakZaglavlja).Value), _
            Range(cfg.get_kolicina & i).Value, Range(cfg.get_ukupniIznos & i).Value, CStr(Split(Range(cfg.get_lv_lu & i).Value, " | ")(0)), _
            CStr(Range(cfg.get_analitickiTM & i).Value), CStr(Range(cfg.get_analitickiMrezniCvor & i).Value), CStr(fyp))
            
             If i = LastRow - 1 Then
                SQLinsertInvoice = SQLinsertInvoice & queries.insertASISTATUS(CStr(msgid))
            End If
        Next i
        
        If control = True Then
            Debug.Print SQLinsertInvoice
            Set rs = CreateObject("ADODB.Recordset")
            rs.Open SQLinsertInvoice, Cn, adOpenStatic
            Set rs = Nothing
            
            insertLog "insert_invoice", _
                "{ reasonCode: " & Split(Range(cfg.get_reasonCodeTekst & cfg.get_reasonCodeRedak).Value, " | ")(0) _
                & ", headerSite: " & Split(Range(cfg.get_lokacija & cfg.get_zaglavlje).Value, " | ")(0) _
                & ", invoiceType: " & Range(cfg.get_tipFakture & cfg.get_zaglavlje).Value _
                & ", customer: " & Split(Range(cfg.get_kupac & cfg.get_zaglavlje).Value, " | ")(0) _
                & ", contract: " & Split(Range(cfg.get_ugovor & cfg.get_zaglavlje).Value, " | ")(0) _
                & ", date: " & Range(cfg.get_datumFakture & cfg.get_zaglavlje).Value _
                & ", remark: " & Range(cfg.get_napomena & cfg.get_zaglavlje).Value _
                & " }", CStr(SQLinsertInvoice)
            
            MsgBox "Faktura u pripremi je upješno poslana GOLD!", vbOKOnly, "Informacija"
            
        Else
            MsgBox "Na popisu u stupcu B se nalaze nedozvoljeni artikli, faktura nije poslana u GOLD!", vbCritical, "Greška"
        End If
                
        
        Cn.Close
        Set Cn = Nothing
        
    ElseIf ans = 7 Then
        'NO
    End If
    
    ' kasnije možemo pokrenuti sa servera program da obradimo insert cijena i nakon toga bi mogli dohvatiti status ažuriranja cijena
    
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub



Sub loadStorno()
    
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    If Len(Range("E2").Value) > 0 And Len(Range("F2").Value) > 0 Then
    
        ans = MsgBox("Jeste li sigurni da želite dohvatiti podatke za storno fakture?", vbYesNo, "Upozorenje")
        If ans = 6 Then
            'YES
            cfg.Init
            Dim sht As Worksheet
            Set sht = ActiveSheet
            Range(cfg.get_korisnik & cfg.get_zaglavlje).Value = utils.getUserName
            Range(cfg.get_lokacija & cfg.get_zaglavlje & ":" & cfg.get_napomena & cfg.get_zaglavlje).ClearContents
            Range(cfg.get_artikl & cfg.get_stavke & ":" & cfg.get_analitickiMrezniCvor & sht.Rows.Count).ClearContents
            Range(cfg.get_lokacija & cfg.get_zaglavlje).Activate
            
            Set Cn = CreateObject("ADODB.Connection")
            Cn.ConnectionTimeout = 1000
            Cn.commandtimeout = 1000
            Cn.Open db.getConnectionString
            
            
            sqlstr = queries.getStornoData(Range("E2").Value, Range("F2").Value)
            'Debug.Print sqlstr
            
            Set rsStorno = CreateObject("ADODB.Recordset")
            rsStorno.Open sqlstr, Cn, adOpenStatic
            
            
            If rsStorno.EOF = False Then
                Dim row As Long
                row = 11
                Do Until rsStorno.EOF = True
                    If row = 11 Then
                        Range("C2").Value = rsStorno(0)
                        Range("C5").Value = rsStorno(1)
                        Range("D5").Value = rsStorno(2)
                        Range("E5").Value = rsStorno(3)
                        Range("F5").Value = rsStorno(4)
                        Range("H5").Value = "STORNO - " & Range("E2").Value & ", " & Range("F2").Value
                        Range("G5").Value = Date
                    End If
                    
                    Range("B" & row).Value = rsStorno(6)
                    Range("C" & row).Value = rsStorno(7)
                    Range("D" & row).Value = rsStorno(8)
                    Range("E" & row).Value = rsStorno(9)
                    Range("F" & row).Value = rsStorno(10)
                    Range("G" & row).Value = rsStorno(11)
                    Range("H" & row).Value = rsStorno(12)
                    Range("I" & row).Value = rsStorno(13)
                    Range("J" & row).Value = rsStorno(14)
            
                    row = row + 1
                  
                    
                    rsStorno.MoveNext
                Loop
                Application.ScreenUpdating = True
                Application.Cursor = xlDefault
                MsgBox "Storno fakture je uspješno pripremljen!", vbOKOnly, "Informacija"
            Else
                MsgBox "Faktrua nije napravljena putem excel alata! Potreban storno kroz GOLD ekran!", vbOKOnly, "Informacija"
            End If

            
            rsStorno.Close
            Set rsStorno = Nothing
            
            insertLog "get_storno", _
            "{ originalInvNum: " & Range("E2").Value _
            & ", originalInvDate: " & Range("F2").Value _
            & " }", CStr(sqlstr)
              
        ElseIf ans = 7 Then
            'NO
        End If
        
    Else
        MsgBox "Broj i datum fakture su obavezna polja!", vbOKOnly, "Informacija"
        Range("E2").Select
    End If
    
    ' kasnije možemo pokrenuti sa servera program da obradimo insert cijena i nakon toga bi mogli dohvatiti status ažuriranja cijena
    
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault

End Sub

