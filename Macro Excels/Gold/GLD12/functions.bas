Attribute VB_Name = "functions"
Sub insertLog(operation As String, parameters As String, sqlquery As String)
    Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        Sqlstr = queries.getLog(db.getDocType, db.getDocName, db.getDocVersion, utils.getUserName, operation, parameters, Replace(sqlquery, "'", """"))
        'Debug.Print SQLstr
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open Sqlstr, Cn, adOpenStatic
        
        Cn.Close
        Set Cn = Nothing
End Sub


Sub loadOrders(msgid As String)
    cfg.Init
    
    If Not IsEmpty(Range("C7").Value) Then
        site = Range("C7").Value
    End If
    
    If Not IsEmpty(Range("C9").Value) Then
        deliveryDate = Range("C9").Value
    End If
    
    stores = "-1"
    If Not IsEmpty(Range("C11").Value) Then
        stores = Range("C11").Value
    End If
    
    barcodes = "-1"
    If Not IsEmpty(Range("E6:E" & utils.getLastRow("E")).Value) Then
        barcodes = ""
        For i = 6 To utils.getLastRow("E")
            If (Len(Range("E" & i).Value) > 0) Then
                If (i = utils.getLastRow("E") - 1) Then
                    barcodes = barcodes & "''" & Range("E" & i).Value & "''"
                Else
                    barcodes = barcodes & "''" & Range("E" & i).Value & "'',"
                End If
            End If
            
        Next i
    End If
    
    
    If Len(Range("C7").Value) = 0 Then
        MsgBox "Potrebno je upisati šifru skladišta!", vbOKOnly, "Greška"
        Range("C7").Activate
        globals.setAllowEventHandling True
        Application.ScreenUpdating = True
        Application.Cursor = xlDefault
        Exit Sub
    End If
    
    
    If Not IsDate(Range("C9").Value) Then
        MsgBox "Planirani datum isporuke je obavezno polje!", vbOKOnly, "Greška"
        Range("C9").Activate
        globals.setAllowEventHandling True
        Application.ScreenUpdating = True
        Application.Cursor = xlDefault
        Exit Sub
    End If
    
    
    If barcodes = "-1" Then
        MsgBox "Potrebno je upisati barkodove!", vbOKOnly, "Greška"
        Range("E6").Activate
        globals.setAllowEventHandling True
        Application.ScreenUpdating = True
        Application.Cursor = xlDefault
        Exit Sub
    End If
  
    
    Sheets(2).Select
    Range(cfg.getcolEAN & "5:" & cfg.getcolINTPROPER & utils.getLastRow(cfg.getcolEAN)).ClearContents
    Application.Goto Range("E5"), True

    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    If Len(msgid) > 0 Then
        Sqlstr = queries.loadDBResponse(CStr(msgid))
    Else
        Sqlstr = queries.loadOrders(CStr(site), CStr(utils.getDateString(CDate(deliveryDate))), CStr(barcodes), CStr(stores))
    End If
    
    'Debug.Print (Sqlstr)
    
    insertLog "load_orders", _
    "{ date: " & Date _
    & ", siteFrom: " & site _
    & ", barcodes: [" & barcodes & "]" _
    & ", sitesTo: [" & stores & "]" _
    & ", deliveryDate: " & utils.getDateString(CDate(deliveryDate)) _
    & " }", CStr(Sqlstr)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open Sqlstr, Cn, adOpenStatic
    
    If rs.EOF = False Then
    Dim row As Long
    row = 5
    Do Until rs.EOF = True
        
        Range(cfg.getcolEAN & row).Value = rs(cfg.getrsEAN)
        Range(cfg.getcolNAZIV & row).Value = rs(cfg.getrsNAZIV)
        Range(cfg.getcolINTSITE & row).Value = rs(cfg.getrsINTSITE)
        Range(cfg.getcolINTQTEC & row).Value = rs(cfg.getrsINTQTEC)
        Range(cfg.getcolINTDCOM & row).Value = rs(cfg.getrsINTDCOM)
        Range(cfg.getcolINTDLIV & row).Value = rs(cfg.getrsINTDLIV)
        Range(cfg.getcolINTID & row).Value = rs(cfg.getrsINTID)
        Range(cfg.getcolINTLCDE & row).Value = rs(cfg.getrsINTLCDE)
        Range(cfg.getcolINTCNUF & row).Value = rs(cfg.getrsINTCNUF)
        Range(cfg.getcolINTCCOM & row).Value = rs(cfg.getrsINTCCOM)
        Range(cfg.getcolINTNFILF & row).Value = rs(cfg.getrsINTNFILF)
        Range(cfg.getcolINTFILC & row).Value = rs(cfg.getrsINTFILC)
        Range(cfg.getcolINTCONF & row).Value = rs(cfg.getrsINTCONF)
        Range(cfg.getcolINTGREL & row).Value = rs(cfg.getrsINTGREL)
        Range(cfg.getcolINTDEVI & row).Value = rs(cfg.getrsINTDEVI)
        Range(cfg.getcolINTCOUC & row).Value = rs(cfg.getrsINTCOUC)
        Range(cfg.getcolINTTXCH & row).Value = rs(cfg.getrsINTTXCH)
        Range(cfg.getcolINTCOM1 & row).Value = rs(cfg.getrsINTCOM1)
        Range(cfg.getcolINTCOM2 & row).Value = rs(cfg.getrsINTCOM2)
        Range(cfg.getcolINTENLEV & row).Value = rs(cfg.getrsINTENLEV)
        Range(cfg.getcolINTDLIM & row).Value = rs(cfg.getrsINTDLIM)
        Range(cfg.getcolINTCODE & row).Value = rs(cfg.getrsINTCODE)
        Range(cfg.getcolINTRCOM & row).Value = rs(cfg.getrsINTRCOM)
        Range(cfg.getcolINTCEXVA & row).Value = rs(cfg.getrsINTCEXVA)
        Range(cfg.getcolINTCEXVL & row).Value = rs(cfg.getrsINTCEXVL)
        Range(cfg.getcolINTUAUVC & row).Value = rs(cfg.getrsINTUAUVC)
        Range(cfg.getcolINTNEGO & row).Value = rs(cfg.getrsINTNEGO)
        Range(cfg.getcolINTORDR & row).Value = rs(cfg.getrsINTORDR)
        Range(cfg.getcolINTSTAT & row).Value = rs(cfg.getrsINTSTAT)
        Range(cfg.getcolINTCEXGLO & row).Value = rs(cfg.getrsINTCEXGLO)
        Range(cfg.getcolINTNOOE & row).Value = rs(cfg.getrsINTNOOE)
        Range(cfg.getcolINTFLUX & row).Value = rs(cfg.getrsINTFLUX)
        Range(cfg.getcolINTFSTA & row).Value = rs(cfg.getrsINTFSTA)
        Range(cfg.getcolINTLDIST & row).Value = rs(cfg.getrsINTLDIST)
        Range(cfg.getcolINTLDNO & row).Value = rs(cfg.getrsINTLDNO)
        Range(cfg.getcolINTETAT & row).Value = rs(cfg.getrsINTETAT)
        Range(cfg.getcolINTSITLI & row).Value = rs(cfg.getrsINTSITLI)
        Range(cfg.getcolINTPACH & row).Value = rs(cfg.getrsINTPACH)
        Range(cfg.getcolINTCOML1 & row).Value = rs(cfg.getrsINTCOML1)
        Range(cfg.getcolINTCLCUS & row).Value = rs(cfg.getrsINTCLCUS)
        Range(cfg.getcolINTURG & row).Value = rs(cfg.getrsINTURG)
        Range(cfg.getcolINTEXT & row).Value = rs(cfg.getrsINTEXT)
        Range(cfg.getcolINTESCO & row).Value = rs(cfg.getrsINTESCO)
        Range(cfg.getcolINTNJESC & row).Value = rs(cfg.getrsINTNJESC)
        Range(cfg.getcolINTPORI & row).Value = rs(cfg.getrsINTPORI)
        Range(cfg.getcolINTINCO & row).Value = rs(cfg.getrsINTINCO)
        Range(cfg.getcolINTLIEU2 & row).Value = rs(cfg.getrsINTLIEU2)
        Range(cfg.getcolINTTRSP & row).Value = rs(cfg.getrsINTTRSP)
        Range(cfg.getcolINTFRAN & row).Value = rs(cfg.getrsINTFRAN)
        Range(cfg.getcolINTVOLI & row).Value = rs(cfg.getrsINTVOLI)
        Range(cfg.getcolINTPDSI & row).Value = rs(cfg.getrsINTPDSI)
        Range(cfg.getcolINTTYIM & row).Value = rs(cfg.getrsINTTYIM)
        Range(cfg.getcolINTDBAS & row).Value = rs(cfg.getrsINTDBAS)
        Range(cfg.getcolINTDDEP & row).Value = rs(cfg.getrsINTDDEP)
        Range(cfg.getcolINTCRED & row).Value = rs(cfg.getrsINTCRED)
        Range(cfg.getcolINTJOUR & row).Value = rs(cfg.getrsINTJOUR)
        Range(cfg.getcolINTDARR & row).Value = rs(cfg.getrsINTDARR)
        Range(cfg.getcolINTMREG & row).Value = rs(cfg.getrsINTMREG)
        Range(cfg.getcolINTDDS & row).Value = rs(cfg.getrsINTDDS)
        Range(cfg.getcolINTNBJM & row).Value = rs(cfg.getrsINTNBJM)
        Range(cfg.getcolINTDVAL & row).Value = rs(cfg.getrsINTDVAL)
        Range(cfg.getcolINTDPAI & row).Value = rs(cfg.getrsINTDPAI)
        Range(cfg.getcolINTNSEQ & row).Value = rs(cfg.getrsINTNSEQ)
        Range(cfg.getcolINTNLIG & row).Value = rs(cfg.getrsINTNLIG)
        Range(cfg.getcolINTNLEN & row).Value = rs(cfg.getrsINTNLEN)
        Range(cfg.getcolINTFICH & row).Value = rs(cfg.getrsINTFICH)
        Range(cfg.getcolINTCACT & row).Value = rs(cfg.getrsINTCACT)
        Range(cfg.getcolINTNERR & row).Value = rs(cfg.getrsINTNERR)
        Range(cfg.getcolINTMESS & row).Value = rs(cfg.getrsINTMESS)
        Range(cfg.getcolINTDCRE & row).Value = rs(cfg.getrsINTDCRE)
        Range(cfg.getcolINTDMAJ & row).Value = rs(cfg.getrsINTDMAJ)
        Range(cfg.getcolINTUTIL & row).Value = rs(cfg.getrsINTUTIL)
        Range(cfg.getcolINTDTRT & row).Value = rs(cfg.getrsINTDTRT)
        Range(cfg.getcolINTCTVA & row).Value = rs(cfg.getrsINTCTVA)
        Range(cfg.getcolINTUAPP & row).Value = rs(cfg.getrsINTUAPP)
        Range(cfg.getcolINTALTF & row).Value = rs(cfg.getrsINTALTF)
        Range(cfg.getcolINTTYPUL & row).Value = rs(cfg.getrsINTTYPUL)
        Range(cfg.getcolINTCEXOGL & row).Value = rs(cfg.getrsINTCEXOGL)
        Range(cfg.getcolINTCEXOPS & row).Value = rs(cfg.getrsINTCEXOPS)
        Range(cfg.getcolINTNROUTE & row).Value = rs(cfg.getrsINTNROUTE)
        Range(cfg.getcolINTLIEU & row).Value = rs(cfg.getrsINTLIEU)
        Range(cfg.getcolINTVALOF & row).Value = rs(cfg.getrsINTVALOF)
        Range(cfg.getcolINTMOTIF & row).Value = rs(cfg.getrsINTMOTIF)
        Range(cfg.getcolINTTEL & row).Value = rs(cfg.getrsINTTEL)
        Range(cfg.getcolINTORI & row).Value = rs(cfg.getrsINTORI)
        Range(cfg.getcolINTCSIN & row).Value = rs(cfg.getrsINTCSIN)
        Range(cfg.getcolINTCTLA & row).Value = rs(cfg.getrsINTCTLA)
        Range(cfg.getcolINTIRECYC & row).Value = rs(cfg.getrsINTIRECYC)
        Range(cfg.getcolINTCRGP & row).Value = rs(cfg.getrsINTCRGP)
        Range(cfg.getcolINTFLIR & row).Value = rs(cfg.getrsINTFLIR)
        Range(cfg.getcolINTNOLV & row).Value = rs(cfg.getrsINTNOLV)
        Range(cfg.getcolINTDRAM & row).Value = rs(cfg.getrsINTDRAM)
        Range(cfg.getcolINTPVSA & row).Value = rs(cfg.getrsINTPVSA)
        Range(cfg.getcolINTPVSR & row).Value = rs(cfg.getrsINTPVSR)
        Range(cfg.getcolINTPRFA & row).Value = rs(cfg.getrsINTPRFA)
        Range(cfg.getcolINTMTDR & row).Value = rs(cfg.getrsINTMTDR)
        Range(cfg.getcolINTMTVI & row).Value = rs(cfg.getrsINTMTVI)
        Range(cfg.getcolINTGRA & row).Value = rs(cfg.getrsINTGRA)
        Range(cfg.getcolINTDENVREC & row).Value = rs(cfg.getrsINTDENVREC)
        Range(cfg.getcolINTCEAN & row).Value = rs(cfg.getrsINTCEAN)
        Range(cfg.getcolINTCEXTJF & row).Value = rs(cfg.getrsINTCEXTJF)
        Range(cfg.getcolINTEDOU & row).Value = rs(cfg.getrsINTEDOU)
        Range(cfg.getcolINTRDOU & row).Value = rs(cfg.getrsINTRDOU)
        Range(cfg.getcolINTDENLEV & row).Value = rs(cfg.getrsINTDENLEV)
        Range(cfg.getcolINTREFEXT & row).Value = rs(cfg.getrsINTREFEXT)
        Range(cfg.getcolINTCTRL & row).Value = rs(cfg.getrsINTCTRL)
        Range(cfg.getcolINTFVSA & row).Value = rs(cfg.getrsINTFVSA)
        Range(cfg.getcolINTFVSR & row).Value = rs(cfg.getrsINTFVSR)
        Range(cfg.getcolINTCODLOG & row).Value = rs(cfg.getrsINTCODLOG)
        Range(cfg.getcolINTCODCAI & row).Value = rs(cfg.getrsINTCODCAI)
        Range(cfg.getcolINTUEREMP & row).Value = rs(cfg.getrsINTUEREMP)
        Range(cfg.getcolINTCINB & row).Value = rs(cfg.getrsINTCINB)
        Range(cfg.getcolINTNOLIGN & row).Value = rs(cfg.getrsINTNOLIGN)
        Range(cfg.getcolINTPROPER & row).Value = rs(cfg.getrsINTPROPER)

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


Sub insertOrders()
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    
    If Application.WorksheetFunction.Sum(Range("AD5:AD100000")) > 0 Then
        MsgBox "Narudbženice su veæ ubaèene u GOLD!", vbExclamation, "Upozorenje"
        Exit Sub
    End If
    

ans = MsgBox("Jeste li sigurni da želite spremiti narudžbenice?", vbYesNo, "Upozorenje")
    
    If ans = 6 Then
        'YES
        cfg.Init
        
        Dim i As Long
        lastRow = utils.getLastRow(cfg.getcolEAN)
        
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        SqlMSGID = queries.selectMSGID
        Set rsMSGID = CreateObject("ADODB.Recordset")
        rsMSGID.Open SqlMSGID, Cn, adOpenStatic
        msgid = rsMSGID(0)
        rsMSGID.Close
        Set rsMSGID = Nothing
        
        SqlSEQ = queries.selectSEQ
        Set rsSEQ = CreateObject("ADODB.Recordset")
        rsSEQ.Open SqlSEQ, Cn, adOpenStatic
        seq = rsSEQ(0)
        rsSEQ.Close
        Set rsSEQ = Nothing
        
        SQLorders = ""
        orders = ""
        For i = 5 To lastRow - 1
            If (Range(cfg.getcolINTQTEC & i).Value > 0) Then
                SQLorders = SQLorders + queries.insertOrder(i, CStr(msgid), CStr(seq))
            
                orders = orders + "{ ean: " + CStr(Range(cfg.getcolEAN & i).Value) + "" _
                    + ", naziv: " + Replace(Range(cfg.getcolNAZIV & i).Value, "'", " ") + "" _
                    + ", trgovina: " + CStr(Range(cfg.getcolINTSITE & i).Value) + "" _
                    + ", kolièina: " + CStr(Range(cfg.getcolINTQTEC & i).Value) + " },"
            End If
        Next i
                
                
        'Debug.Print SQLorders
        
        Set rsOrders = CreateObject("ADODB.Recordset")
        rsOrders.Open SQLorders, Cn, adOpenStatic
        Set rsOrders = Nothing
        
        insertLog "insert_orders", _
        "{ site: " & Sheets(1).Range("C7").Value & _
        ", deliveryDate: " & Sheets(1).Range("C9").Value & _
        ", orders: [" & orders & "]" _
        & " }", CStr(SQLorders)

        
        ssh.procesOrders Sheets(1).Range("C11").Value
        
        SQLcheck = queries.checkNonProccesed
        'Debug.Print SQLcheck
        Set rsCheck = CreateObject("ADODB.Recordset")
        rsCheck.Open SQLcheck, Cn, adOpenStatic
        
        If rsCheck.EOF = False Then
            MsgBox "Narudžbenice NISU poslane u GOLD!", vbCritical, "Upozorenje"
            SQLdelete = queries.deleteNonProccesed
            'Debug.Print SQLdelete
            Set rsDelete = CreateObject("ADODB.Recordset")
            rsDelete.Open SQLdelete, Cn, adOpenStatic
            Set rsDelete = Nothing
        Else
            MsgBox "Narudbženice su poslane u GOLD!", vbOKOnly, "Informacija"
            Sheets(2).Select
            Range(cfg.getcolEAN & "5:" & cfg.getcolINTPROPER & utils.getLastRow(cfg.getcolEAN)).ClearContents
            Application.Goto Range("E5"), True
            Sheets(1).Select
            functions.loadOrders CStr(msgid)
            
        End If
        Set rsCheck = Nothing
        
        
        Cn.Close
        Set Cn = Nothing
    ElseIf ans = 7 Then
        'NO
    End If
    
    ' kasnije možemo pokrenuti sa servera program da obradimo insert cijena i nakon toga bi mogli dohvatiti status ažuriranja cijena
   
    
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

