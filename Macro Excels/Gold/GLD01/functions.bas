Attribute VB_Name = "functions"
Sub insertLog(operation As String, parameters As String, sqlquery As String)
    Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        SQLstr = queries.getLog(db.getDocType, db.getDocName, db.getDocVersion, utils.getUserName, operation, parameters, Replace(sqlquery, "'", """"))
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLstr, Cn, adOpenStatic
        
        Cn.Close
        Set Cn = Nothing
End Sub


Sub loadSearch()
    frmSearch.Show
End Sub


Sub loadPurchaseConditions()

    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    globals.setAllowEventHandling False
    
    If Len(Range("C7").Value) > 0 Then
        
        'ulazni parametri sa prvog sheeta
        Dim num_log As Integer
        num_log = 123
        Dim text As String
        text = "-1"
        Dim num As Integer
        num = -1
        
        Dim domain_user As String
        domain_user = Range("C5").Value
        
        Dim indate As String
        indate = Application.WorksheetFunction.text(Range("C7").Value, "dd/mm/yyyy")
        
        Dim site As Integer
        site = -1
        If Len(Range("C9").Value) > 0 Then
            site = CInt(Split(Range("C9").Value, " - ")(0))
        End If
        
        Dim cnuf As String
        cnuf = "-1"
        If Len(Range("C11").Value) > 0 Then
            cnuf = Split(Range("C11").Value, " - ")(0)
        End If
    
        Dim cnum As String
        cnum = "-1"
        If Len(Range("C13").Value) > 0 Then
            cnum = Split(Range("C13").Value, " - ")(0)
        End If
        
        Dim ms As String
        ms = -1
        If Len(Range("C15").Value) > 0 Then
            ms = Split(Range("C15").Value, " - ")(0)
        End If
        
        
        Dim artlist As String
        artlist = -1
        If Len(Range("C17").Value) > 0 Then
            artlist = Split(Range("C17").Value, " - ")(0)
        End If

        Dim art_grp As String
        art_grp = "-1"
        If Len(Range("C18").Value) > 0 Then
            art_grp = Split(Range("C18").Value, " - ")(0)
        End If
        
        Dim cexr As String
        cexr = -1
        If Len(Range("C19").Value) > 0 Then
            cexr = Split(Range("C19").Value, " - ")(0)
        End If
        
  
        Dim ccla As String
        ccla = -1
        If Len(Range("C21").Value) > 0 Then
            art_grp = Split(Range("C21").Value, " - ")(0)
        End If
        Dim catt As String
        catt = -1
        If Len(Range("C22").Value) > 0 Then
            art_grp = Split(Range("C22").Value, " - ")(0)
        End If
        
        globals.setOldConditions (Range("F13").Value)
        globals.setFutureConditions (Range("F14").Value)
        
        barcodes = "''-1''"
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
                       
        
        'punjenje interface tablice
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        SQLstr = queries.loadPurchaseConditionsDataToInterfaceTable(num_log, cnuf, cnum, indate, art_grp, cexr, site, ms, ccla, catt, text, num, artlist, domain_user)
        
        insertLog "load_purchase_conditions", _
        "{ date: " & Range("C7").Value _
        & ", location: " & Range("C9").Value _
        & ", supplier: " & Range("C11").Value _
        & ", contract: " & Range("C13").Value _
        & ", pastConditions: " & Range("F13").Value _
        & ", futConditions: " & Range("F14").Value _
        & ", ms: " & Range("C15").Value _
        & ", articleList: " & Range("C17").Value _
        & ", articleGroup: " & Range("C18").Value _
        & ", article: " & cexr _
        & ", barcodes: [" & barcodes & "]" _
        & ", class: " & Range("C21").Value _
        & ", classAttribute: " & Range("C22").Value _
        & " }", CStr(SQLstr)
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLstr, Cn, adOpenStatic
        
        Dim msgid As String
        msgid = ""
        mess = rs(0)
        num_lines = rs(1)
        
        cfg.InitCollections
        
        If num_lines > 0 Then
            msgid = rs(2)
            'dohvat podataka iz interface tablice
            cfg.Init
            
            ActiveWorkbook.Sheets(3).Select
            Range(cfg.getcTNUMSGID & "6:" & cfg.getcBROJPROMJENA & utils.getLastRow(cfg.getcTNUMSGID)).ClearContents
            
            ActiveWorkbook.Sheets(2).Select
            Range(cfg.getcTNUMSGID & "6:" & cfg.getcBROJPROMJENA & utils.getLastRow(cfg.getcTNUMSGID)).Select
            Selection.ClearContents
            Selection.ClearComments
            With Selection.Font
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
            End With
            
            
            ActiveWorkbook.Sheets(4).Visible = xlSheetVisible
            ActiveWorkbook.Sheets(4).Range("A2:F99999").ClearContents
            ActiveWorkbook.Sheets(4).Visible = xlSheetVeryHidden
            
            Set Cn = CreateObject("ADODB.Connection")
            Cn.ConnectionTimeout = 1000
            Cn.commandtimeout = 1000
            Cn.Open db.getConnectionString
            
            SQLstr = queries.selectPurchaseConditionsDataFromInterfaceTable(msgid, CStr(barcodes))
            
            'Debug.Print SQLStr
            
            Set rs = CreateObject("ADODB.Recordset")
            rs.Open SQLstr, Cn, adOpenStatic
            
            row = 6
            Do Until rs.EOF = True
                Range(cfg.getcTNUMSGID & row).Value = rs(cfg.getrTNUMSGID)
                Range(cfg.getcTNULNLIG & row).Value = rs(cfg.getrTNULNLIG)
                Range(cfg.getcTNUCNUF & row).Value = rs(cfg.getrTNUCNUF)
                Range(cfg.getcTNUSUPDESC & row).Value = rs(cfg.getrTNUSUPDESC)
                Range(cfg.getcTNUCCOM & row).Value = rs(cfg.getrTNUCCOM)
                Range(cfg.getcTNUAGRP & row).Value = rs(cfg.getrTNUAGRP)
                Range(cfg.getcTNUCEXR & row).Value = rs(cfg.getrTNUCEXR)
                Range(cfg.getcARCCODE & row).Value = rs(cfg.getrARCCODE)
                Range(cfg.getcTNUADESC & row).Value = rs(cfg.getrTNUADESC)
                Range(cfg.getcTNULV & row).Value = rs(cfg.getrTNULV)
                Range(cfg.getcTNULU & row).Value = rs(cfg.getrTNULU)
                Range(cfg.getcTNUSITE & row).Value = rs(cfg.getrTNUSITE)
                Range(cfg.getcTNUSDESC & row).Value = rs(cfg.getrTNUSDESC)
                Range(cfg.getcPRINCIPAL & row).Value = rs(cfg.getrPRINCIPAL)
                Range(cfg.getcASORTIMAN & row).Value = rs(cfg.getrASORTIMAN)
                Range(cfg.getcTNUPACH & row).Value = rs(cfg.getrTNUPACH)
                utils.addComment cfg.getcTNUPACH & row, rs(cfg.getrTNUPASTPACH), rs(cfg.getrTNUFUTPACH)
                Range(cfg.getcTNUUAPP & row).Value = rs(cfg.getrTNUUAPP)
                Range(cfg.getcTNUNNC & row).Value = rs(cfg.getrTNUNNC)
                Range(cfg.getcTNUEXNNC & row).Value = rs(cfg.getrTNUEXNNC)
                Range(cfg.getcTNUPADDEB & row).Value = CDate(rs(cfg.getrTNUPADDEB))
                Range(cfg.getcTNUPADFIN & row).Value = CDate(rs(cfg.getrTNUPADFIN))
                Range(cfg.getcTNUTCP & row).Value = rs(cfg.getrTNUTCP)
                
                If rs(cfg.getrTNUVAL601) > 0 Then
                    Range(cfg.getcTNUVAL601 & row).Value = rs(cfg.getrTNUVAL601)
                    Range(cfg.getcTNUUAPP601 & row).Value = rs(cfg.getrTNUUAPP601)
                    Range(cfg.getcTNUDDEB601 & row).Value = CDate(rs(cfg.getrTNUDDEB601))
                    Range(cfg.getcTNUDFIN601 & row).Value = CDate(rs(cfg.getrTNUDFIN601))
                End If
                utils.addComment cfg.getcTNUVAL601 & row, rs(cfg.getrTNUPAST601), rs(cfg.getrTNUFUT601)
                
                If rs(cfg.getrTNUVAL602) > 0 Then
                    Range(cfg.getcTNUVAL602 & row).Value = rs(cfg.getrTNUVAL602)
                    Range(cfg.getcTNUUAPP602 & row).Value = rs(cfg.getrTNUUAPP602)
                    Range(cfg.getcTNUDDEB602 & row).Value = CDate(rs(cfg.getrTNUDDEB602))
                    Range(cfg.getcTNUDFIN602 & row).Value = CDate(rs(cfg.getrTNUDFIN602))
                End If
                utils.addComment cfg.getcTNUVAL602 & row, rs(cfg.getrTNUPAST602), rs(cfg.getrTNUFUT602)
                
                If rs(cfg.getrTNUVAL603) > 0 Then
                    Range(cfg.getcTNUVAL603 & row).Value = rs(cfg.getrTNUVAL603)
                    Range(cfg.getcTNUUAPP603 & row).Value = rs(cfg.getrTNUUAPP603)
                    Range(cfg.getcTNUDDEB603 & row).Value = CDate(rs(cfg.getrTNUDDEB603))
                    Range(cfg.getcTNUDFIN603 & row).Value = CDate(rs(cfg.getrTNUDFIN603))
                End If
                utils.addComment cfg.getcTNUVAL603 & row, rs(cfg.getrTNUPAST603), rs(cfg.getrTNUFUT603)
                
                If rs(cfg.getrTNUVAL604) > 0 Then
                    Range(cfg.getcTNUVAL604 & row).Value = rs(cfg.getrTNUVAL604)
                    Range(cfg.getcTNUUAPP604 & row).Value = rs(cfg.getrTNUUAPP604)
                    Range(cfg.getcTNUDDEB604 & row).Value = CDate(rs(cfg.getrTNUDDEB604))
                    Range(cfg.getcTNUDFIN604 & row).Value = CDate(rs(cfg.getrTNUDFIN604))
                End If
                utils.addComment cfg.getcTNUVAL604 & row, rs(cfg.getrTNUPAST604), rs(cfg.getrTNUFUT604)
                
                If rs(cfg.getrTNUVAL605) > 0 Then
                    Range(cfg.getcTNUVAL605 & row).Value = rs(cfg.getrTNUVAL605)
                    Range(cfg.getcTNUUAPP605 & row).Value = rs(cfg.getrTNUUAPP605)
                    Range(cfg.getcTNUDDEB605 & row).Value = CDate(rs(cfg.getrTNUDDEB605))
                    Range(cfg.getcTNUDFIN605 & row).Value = CDate(rs(cfg.getrTNUDFIN605))
                End If
                utils.addComment cfg.getcTNUVAL605 & row, rs(cfg.getrTNUPAST605), rs(cfg.getrTNUFUT605)
                
                If rs(cfg.getrTNUVAL606) > 0 Then
                    Range(cfg.getcTNUVAL606 & row).Value = rs(cfg.getrTNUVAL606)
                    Range(cfg.getcTNUUAPP606 & row).Value = rs(cfg.getrTNUUAPP606)
                    Range(cfg.getcTNUDDEB606 & row).Value = CDate(rs(cfg.getrTNUDDEB606))
                    Range(cfg.getcTNUDFIN606 & row).Value = CDate(rs(cfg.getrTNUDFIN606))
                End If
                utils.addComment cfg.getcTNUVAL606 & row, rs(cfg.getrTNUPAST606), rs(cfg.getrTNUFUT606)
                
                
                cfg.addKeyItem CStr("$" & cfg.getcTNUPACH & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUUAPP & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUNNC & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUEXNNC & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUPADDEB & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUPADFIN & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUTCP & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUVAL601 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUUAPP601 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUDDEB601 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUDFIN601 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUVAL602 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUUAPP602 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUDDEB602 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUDFIN602 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUVAL603 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUUAPP603 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUDDEB603 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUDFIN603 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUVAL604 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUUAPP604 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUDDEB604 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUDFIN604 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUVAL605 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUUAPP605 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUDDEB605 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUDFIN605 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUVAL606 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUUAPP606 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUDDEB606 & "$" & row)
                cfg.addKeyItem CStr("$" & cfg.getcTNUDFIN606 & "$" & row)
                
                
                cfg.addKeyValue CStr("$" & cfg.getcTNUPACH & "$" & row), utils.getString(rs(cfg.getrTNUPACH), 1, 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUUAPP & "$" & row), utils.getString(rs(cfg.getrTNUUAPP), 1, 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUNNC & "$" & row), utils.getString(rs(cfg.getrTNUNNC), 1, 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUEXNNC & "$" & row), utils.getString(rs(cfg.getrTNUEXNNC), 1, 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUPADDEB & "$" & row), utils.getString(rs(cfg.getrTNUPADDEB), 1, 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUPADFIN & "$" & row), utils.getString(rs(cfg.getrTNUPADFIN), 1, 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUTCP & "$" & row), utils.getString(rs(cfg.getrTNUTCP), 1, 0)
                
   
                cfg.addKeyValue CStr("$" & cfg.getcTNUVAL601 & "$" & row), utils.getString(rs(cfg.getrTNUVAL601), rs(cfg.getrTNUVAL601), 1)
                cfg.addKeyValue CStr("$" & cfg.getcTNUUAPP601 & "$" & row), utils.getString(rs(cfg.getrTNUUAPP601), rs(cfg.getrTNUVAL601), 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUDDEB601 & "$" & row), utils.getString(rs(cfg.getrTNUDDEB601), rs(cfg.getrTNUVAL601), 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUDFIN601 & "$" & row), utils.getString(rs(cfg.getrTNUDFIN601), rs(cfg.getrTNUVAL601), 0)
       
                
                cfg.addKeyValue CStr("$" & cfg.getcTNUVAL602 & "$" & row), utils.getString(rs(cfg.getrTNUVAL602), rs(cfg.getrTNUVAL602), 1)
                cfg.addKeyValue CStr("$" & cfg.getcTNUUAPP602 & "$" & row), utils.getString(rs(cfg.getrTNUUAPP602), rs(cfg.getrTNUVAL602), 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUDDEB602 & "$" & row), utils.getString(rs(cfg.getrTNUDDEB602), rs(cfg.getrTNUVAL602), 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUDFIN602 & "$" & row), utils.getString(rs(cfg.getrTNUDFIN602), rs(cfg.getrTNUVAL602), 0)
                
                cfg.addKeyValue CStr("$" & cfg.getcTNUVAL603 & "$" & row), utils.getString(rs(cfg.getrTNUVAL603), rs(cfg.getrTNUVAL603), 1)
                cfg.addKeyValue CStr("$" & cfg.getcTNUUAPP603 & "$" & row), utils.getString(rs(cfg.getrTNUUAPP603), rs(cfg.getrTNUVAL603), 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUDDEB603 & "$" & row), utils.getString(rs(cfg.getrTNUDDEB603), rs(cfg.getrTNUVAL603), 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUDFIN603 & "$" & row), utils.getString(rs(cfg.getrTNUDFIN603), rs(cfg.getrTNUVAL603), 0)
                
                cfg.addKeyValue CStr("$" & cfg.getcTNUVAL604 & "$" & row), utils.getString(rs(cfg.getrTNUVAL604), rs(cfg.getrTNUVAL604), 1)
                cfg.addKeyValue CStr("$" & cfg.getcTNUUAPP604 & "$" & row), utils.getString(rs(cfg.getrTNUUAPP604), rs(cfg.getrTNUVAL604), 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUDDEB604 & "$" & row), utils.getString(rs(cfg.getrTNUDDEB604), rs(cfg.getrTNUVAL604), 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUDFIN604 & "$" & row), utils.getString(rs(cfg.getrTNUDFIN604), rs(cfg.getrTNUVAL604), 0)
                
                cfg.addKeyValue CStr("$" & cfg.getcTNUVAL605 & "$" & row), utils.getString(rs(cfg.getrTNUVAL605), rs(cfg.getrTNUVAL605), 1)
                cfg.addKeyValue CStr("$" & cfg.getcTNUUAPP605 & "$" & row), utils.getString(rs(cfg.getrTNUUAPP605), rs(cfg.getrTNUVAL605), 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUDDEB605 & "$" & row), utils.getString(rs(cfg.getrTNUDDEB605), rs(cfg.getrTNUVAL605), 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUDFIN605 & "$" & row), utils.getString(rs(cfg.getrTNUDFIN605), rs(cfg.getrTNUVAL605), 0)
                
                cfg.addKeyValue CStr("$" & cfg.getcTNUVAL606 & "$" & row), utils.getString(rs(cfg.getrTNUVAL606), rs(cfg.getrTNUVAL606), 1)
                cfg.addKeyValue CStr("$" & cfg.getcTNUUAPP606 & "$" & row), utils.getString(rs(cfg.getrTNUUAPP606), rs(cfg.getrTNUVAL606), 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUDDEB606 & "$" & row), utils.getString(rs(cfg.getrTNUDDEB606), rs(cfg.getrTNUVAL606), 0)
                cfg.addKeyValue CStr("$" & cfg.getcTNUDFIN606 & "$" & row), utils.getString(rs(cfg.getrTNUDFIN606), rs(cfg.getrTNUVAL606), 0)
                
                
                row = row + 1
                rs.MoveNext
            Loop
        Else
            MsgBox "Ulazni parametri nisu dali rezultat pretrage!", vbOKOnly, "Informacija"
        End If
        
        rs.Close
        Set rs = Nothing
        Cn.Close
        Set Cn = Nothing
    Else
        MsgBox "Potrebno je upisati datum!", vbOKOnly, "Informacija"
        Range("C7").Select
    End If
    
    ActiveWorkbook.Sheets(4).Visible = xlSheetVisible
    Dim n As Long
    If (cfg.keys.Count) > 0 Then
        For n = 1 To cfg.keys.Count
            ActiveWorkbook.Sheets(4).Range("A" & n + 1).Value = cfg.keys.Item(n)
            ActiveWorkbook.Sheets(4).Range("B" & n + 1).Value = Range(cfg.keys.Item(n)).row
            ActiveWorkbook.Sheets(4).Range("C" & n + 1).Value = CStr(cfg.getValueByKey(cfg.keys.Item(n)))
        Next
    End If
    ActiveWorkbook.Sheets(4).Visible = xlSheetVeryHidden
    
    Range("B6").Select
    globals.setAllowEventHandling True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault

End Sub

Sub loadConditionChanges()
    Sheets(3).Range("B6:AV500000").ClearContents
    Sheets(3).Range("B6:AV500000").ClearComments
    
    Sheets(2).Select
    
    If utils.validatePurchaseConditions = True Then
    
        Dim bpr As String
        bpr = cfg.getcBROJPROMJENA
        ActiveSheet.Range("$B$5:$" & bpr & "$" & utils.getLastRow(bpr)).AutoFilter Field:=47, Criteria1:=">0", Operator:=xlAnd
        Range("B5:" & bpr & utils.getLastRow(bpr)).Select
        Selection.Copy
        
        Sheets(3).Select
        Range("B5").Select
        ActiveSheet.Paste
                
        barcodes = ""
        cexr = ""
        If Not IsEmpty(Range(cfg.getcARCCODE & "6:" & cfg.getcARCCODE & utils.getLastRow(cfg.getcARCCODE)).Value) Then
            For i = 6 To utils.getLastRow(cfg.getcARCCODE)
                If (Len(Range(cfg.getcARCCODE & i).Value) > 0) Then
                    If (i = utils.getLastRow(cfg.getcARCCODE) - 1) Then
                        barcodes = barcodes & "''" & Range(cfg.getcARCCODE & i).Value & "''"
                        cexr = cexr & "''" & Range(cfg.getcTNUCEXR & i).Value & "''"
                    Else
                        barcodes = barcodes & "''" & Range(cfg.getcARCCODE & i).Value & "'',"
                        cexr = cexr & "''" & Range(cfg.getcTNUCEXR & i).Value & "''" & "'',"
                    End If
                End If
            Next i
        End If
                       
        insertLog "load_purchase_changes", _
        "{ cexr: [" & cexr & "]" _
        & ", barcodes: [" & barcodes & "]" _
        & " }", ""
        
        
        Range("K6").Select
        
        Sheets(2).Select
        ActiveSheet.Range("$B$5:$" & bpr & "$" & utils.getLastRow(bpr)).AutoFilter Field:=47
        Application.Goto Range("K6"), True
        Range("K6").Select
        
        Sheets(3).Select
    
    End If
    
End Sub

Sub saveChanges()

    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    ans = MsgBox("Jeste li sigurni da želite spremiti promjene?", vbYesNo, "Upozorenje")
    
    If ans = 6 Then
        'YES
        cfg.Init
        
        Dim i As Long
        LastRow = utils.getLastRow(cfg.getcTNUMSGID)
        
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString

        SQLUpdateConditions = ""
        For i = 6 To LastRow - 1
            SQLUpdateConditions = SQLUpdateConditions & queries.updatePurchaseCondition(Range(cfg.getcTNUPACH & i).Value, Range(cfg.getcTNUUAPP & i).Value, Range(cfg.getcTNUNNC & i).Value, Range(cfg.getcTNUEXNNC & i).Value, Range(cfg.getcTNUPADDEB & i).Value, Range(cfg.getcTNUPADFIN & i).Value, _
                                Range(cfg.getcTNUTCP & i).Value, Range(cfg.getcTNUVAL601 & i).Value, Range(cfg.getcTNUUAPP601 & i).Value, Range(cfg.getcTNUDDEB601 & i).Value, Range(cfg.getcTNUDFIN601 & i).Value, Range(cfg.getcTNUVAL602 & i).Value, _
                                Range(cfg.getcTNUUAPP602 & i).Value, Range(cfg.getcTNUDDEB602 & i).Value, Range(cfg.getcTNUDFIN602 & i).Value, Range(cfg.getcTNUVAL603 & i).Value, Range(cfg.getcTNUUAPP603 & i).Value, Range(cfg.getcTNUDDEB603 & i).Value, _
                                Range(cfg.getcTNUDFIN603 & i).Value, Range(cfg.getcTNUVAL604 & i).Value, Range(cfg.getcTNUUAPP604 & i).Value, Range(cfg.getcTNUDDEB604 & i).Value, Range(cfg.getcTNUDFIN604 & i).Value, Range(cfg.getcTNUVAL605 & i).Value, _
                                Range(cfg.getcTNUUAPP605 & i).Value, Range(cfg.getcTNUDDEB605 & i).Value, Range(cfg.getcTNUDFIN605 & i).Value, Range(cfg.getcTNUVAL606 & i).Value, Range(cfg.getcTNUUAPP606 & i).Value, Range(cfg.getcTNUDDEB606 & i).Value, _
                                Range(cfg.getcTNUDFIN606 & i).Value, CStr(Range(cfg.getcTNUMSGID & i).Value), CStr(Range(cfg.getcTNULNLIG & i).Value))
            
            If i = LastRow - 1 Then
                SQLUpdateConditions = SQLUpdateConditions & queries.insertASISTATUS(Range(cfg.getcTNUMSGID & i).Value)
            End If
        Next i
        
        
        'Debug.Print SQLUpdateConditions
        'Debug.Print "#######"
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLUpdateConditions, Cn, adOpenStatic
        
        barcodes = ""
        cexr = ""
        If Not IsEmpty(Range(cfg.getcARCCODE & "6:" & cfg.getcARCCODE & utils.getLastRow(cfg.getcARCCODE)).Value) Then
            For i = 6 To utils.getLastRow(cfg.getcARCCODE)
                If (Len(Range(cfg.getcARCCODE & i).Value) > 0) Then
                    If (i = utils.getLastRow(cfg.getcARCCODE) - 1) Then
                        barcodes = barcodes & "''" & Range(cfg.getcARCCODE & i).Value & "''"
                        cexr = cexr & "''" & Range(cfg.getcTNUCEXR & i).Value & "''"
                    Else
                        barcodes = barcodes & "''" & Range(cfg.getcARCCODE & i).Value & "'',"
                        cexr = cexr & "''" & Range(cfg.getcTNUCEXR & i).Value & "''" & "'',"
                    End If
                End If
            Next i
        End If
                       
        insertLog "save_purchase_conditions", _
        "{ cexr: [" & cexr & "]" _
        & ", barcodes: [" & barcodes & "]" _
        & " }", CStr(SQLUpdateConditions)
        
        Set rs = Nothing
        Cn.Close
        Set Cn = Nothing
        
        MsgBox "Nabavni uvjeti su uspješno ažurirani u GOLD-u!", vbOKOnly, "Informacija"
        
    ElseIf ans = 7 Then
        'NO
    End If
    
    ' kasnije možemo pokrenuti sa servera program da obradimo insert cijena i nakon toga bi mogli dohvatiti status ažuriranja cijena
    
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault

End Sub


