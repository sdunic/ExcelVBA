Attribute VB_Name = "functions"
Sub insertLog(operation As String, parameters As String, sqlquery As String)
    Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        SQLStr = queries.getLog(db.getDocType, db.getDocName, db.getDocVersion, utils.getUserName, operation, parameters, Replace(sqlquery, "'", """"))
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLStr, Cn, adOpenStatic
        
        Cn.Close
        Set Cn = Nothing
End Sub

Sub loadSearch()
    frmSearch.Show
End Sub

Sub loadReceptions()

    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    globals.setAllowEventHandling False
    
    ActiveWorkbook.Sheets(2).Select
    ActiveSheet.Unprotect
    
    If Len(ActiveWorkbook.Sheets(1).Range("C7").Value) > 0 And Len(ActiveWorkbook.Sheets(1).Range("C8").Value) > 0 And _
    (Len(ActiveWorkbook.Sheets(1).Range("C22").Value) > 0 Or Len(ActiveWorkbook.Sheets(1).Range("C24").Value) > 0) Then
        
        'ulazni parametri sa prvog sheeta
        Dim site As String
        site = ActiveWorkbook.Sheets(1).Range("C7").Value
        Dim receptions As String
        receptions = "''" & Replace(ActiveWorkbook.Sheets(1).Range("C22").Value, ",", "'',''") & "''"
        Dim deliveryNums As String
        deliveryNums = "''" & Replace(ActiveWorkbook.Sheets(1).Range("C24").Value, ",", "'',''") & "''"
        
        Dim domain_user As String
        domain_user = ActiveWorkbook.Sheets(1).Range("C5").Value
        Dim row As Long
        
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        
        Range("L3:N3").ClearContents
        Range("A7:Y" & utils.getLastRow("B")).ClearContents
               
        Set rs = CreateObject("ADODB.Recordset")
        SQLStr = queries.selectReceptions(site, receptions, deliveryNums)
        'Debug.Print SQLStr
        
        rs.Open SQLStr, Cn, adOpenStatic
   
        
        row = 7
        Do Until rs.EOF = True
            If row = 7 Then
                Range("L3").Value = rs(18)
                Range("M3").Value = rs(19)
                Range("N3").Value = rs(20)
            End If
                       
            Range("B" & row).Value = rs(0)
            Range("C" & row).Value = rs(1)
            Range("D" & row).Value = rs(2)
            Range("E" & row).Value = rs(3)
            Range("F" & row).Value = rs(4)
            Range("G" & row).Value = rs(5)
            
            Range("H" & row).Value = 0 'rs(6) 'PN
            
            Range("I" & row).Value = rs(7)
            Range("J" & row).Value = rs(8)
            Range("K" & row).Value = rs(9)
            Range("L" & row).Value = rs(10)
            Range("M" & row).Value = rs(11)
            Range("N" & row).Value = rs(12)
            
            Range("O" & row).Value = rs(13)
            Range("P" & row).Value = rs(13)
            
            Range("Q" & row).Value = rs(14)
            
            Range("R" & row).Value = rs(16)
            Range("S" & row).Value = rs(16)
            
            Range("T" & row).Value = "=ROUND(RC[-1]/7.5345,2)"
            Range("U" & row).Value = "=RC[-6]-RC[-5]"
            Range("V" & row).Value = "=RC[-4]-RC[-3]"
            Range("W" & row).Value = "=RC[-6]-RC[-3]"
            
            Range("X" & row).Value = "=RC[-8]*RC[-5]"
            Range("Y" & row).Value = "=RC[-9]*RC[-5]"

            
            row = row + 1
            rs.MoveNext
        Loop
       
       rs.MoveFirst
       Do Until rs.EOF = True
                             
            If rs(6) = 1 Then
                             
                Range("B" & row).Value = rs(0)
                Range("C" & row).Value = rs(1)
                Range("D" & row).Value = rs(2)
                Range("E" & row).Value = rs(3)
                Range("F" & row).Value = rs(4)
                Range("G" & row).Value = rs(5)
                
                Range("H" & row).Value = 1 'rs(6) 'PN
                
                Range("I" & row).Value = rs(7)
                Range("J" & row).Value = rs(8)
                Range("K" & row).Value = rs(9)
                Range("L" & row).Value = 7
                Range("M" & row).Value = "PDV 0%"
                Range("N" & row).Value = 0
                
                Range("O" & row).Value = rs(13)
                Range("P" & row).Value = rs(13)
                
                Range("Q" & row).Value = 0
                
                Range("R" & row).Value = 0
                Range("S" & row).Value = 0
                
                Range("T" & row).Value = "=ROUND(RC[-1]/7.5345,2)"
                Range("U" & row).Value = "=RC[-6]-RC[-5]"
                Range("V" & row).Value = "=RC[-4]-RC[-3]"
                Range("W" & row).Value = "=RC[-6]-RC[-3]"
                
                Range("X" & row).Value = "=RC[-8]*RC[-5]"
                Range("Y" & row).Value = "=RC[-9]*RC[-5]"
                
                 row = row + 1
            End If
           
            rs.MoveNext
        Loop
       
       
        rs.Close
        Set rs = Nothing
        Cn.Close
        Set Cn = Nothing
        
    Else
        ActiveWorkbook.Sheets(1).Select
        MsgBox "Trgovina, dobavljaè i (dokumenti prijema ili brojevi dostavnice) su obavezni podatci!", vbOKOnly, "Informacija"
        Range("C7").Select
    End If
    
    
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
    False, AllowSorting:=True, AllowFiltering:=True
    ActiveSheet.EnableSelection = xlNoRestrictions
    Range("B5").Select
    globals.setAllowEventHandling True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    
   
End Sub



Sub saveInvoice()

    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    ActiveSheet.Unprotect
    
    Dim site As String
    site = ActiveWorkbook.Sheets(1).Range("C7").Value
    Dim util As String
    util = ActiveWorkbook.Sheets(1).Range("C5").Value
    Dim sup As String
    sup = Split(ActiveWorkbook.Sheets(1).Range("C8").Value, " - ")(0)
    Dim invoiceNum As String
    invoiceNum = ActiveWorkbook.Sheets(1).Range("C10").Value
    Dim invoiceDate As String
    invoiceDate = utils.getDateString(ActiveWorkbook.Sheets(1).Range("C11").Value)
    Dim payDate As String
    payDate = utils.getDateString(ActiveWorkbook.Sheets(1).Range("C12").Value)
    Dim fich As String
    fich = Replace(util, ".", "") & Replace(Replace(Replace(Now, " ", ""), ":", ""), ".", "")
    Dim ccin As String
    ccin = Range("L3").Value
    Dim ccom As String
    ccom = Range("M3").Value
    Dim filf As String
    filf = Range("N3").Value
    Dim cfin As String
    cfin = Split(ActiveWorkbook.Sheets(1).Range("C8").Value, " - ")(1)
    Dim cnuf As String
    cnuf = Split(ActiveWorkbook.Sheets(1).Range("C8").Value, " - ")(0)
    Dim totalAmmount As Double
    totalAmmount = ActiveWorkbook.Sheets(1).Range("E20").Value
    
    If Abs(Range("G4").Value) < Range("I4").Value Then
        ans = MsgBox("Jeste li sigurni da želite spremiti fakturu?", vbYesNo, "Upozorenje")
        
        If ans = 6 Then
            'YES
            
            Dim i As Long
            LastRow = utils.getLastRow("B")
            
            Set Cn = CreateObject("ADODB.Connection")
            Cn.ConnectionTimeout = 1000
            Cn.commandtimeout = 1000
            Cn.Open db.getConnectionString
    
            'intcfinv
            Set rs = CreateObject("ADODB.Recordset")
            sqlInsertHeader = sqlInsertHeader & queries.insertInvoiceHeader(cfin, ccin, cnuf, ccom, invoiceNum, invoiceDate, payDate, 1, filf, util, totalAmmount, fich)
            
            'Debug.Print sqlInsertHeader
            rs.Open sqlInsertHeader, Cn, adOpenStatic
            Set rs = Nothing
    
            'intcfbl
            Dim devNumsByVatRates() As Variant
            ReDim devNumsByVatRates(0)
            Dim devNumByVatRate As String
            Dim vatAmmount() As Double
            ReDim vatAmmount(0)
            Dim netAmmount() As Double
            ReDim netAmmount(0)
            
            For i = 7 To LastRow - 1
                devNumByVatRate = Range("C" & i).Value & "##" & Range("N" & i).Value
                If Not utils.IsInArray(devNumByVatRate, devNumsByVatRates) Then
                    devNumsByVatRates(UBound(devNumsByVatRates, 1) - LBound(devNumsByVatRates, 1)) = devNumByVatRate
                    netAmmount(UBound(netAmmount, 1) - LBound(netAmmount, 1)) = Range("P" & i).Value * Range("T" & i).Value
                    vatAmmount(UBound(vatAmmount, 1) - LBound(vatAmmount, 1)) = Range("P" & i).Value * Range("T" & i).Value * (Range("N" & i).Value / 100)
                    ReDim Preserve devNumsByVatRates(0 To UBound(devNumsByVatRates, 1) - LBound(devNumsByVatRates, 1) + 1)
                    ReDim Preserve vatAmmount(0 To UBound(vatAmmount, 1) - LBound(vatAmmount, 1) + 1)
                    ReDim Preserve netAmmount(0 To UBound(netAmmount, 1) - LBound(netAmmount, 1) + 1)
                Else
                    ind = utils.GetArrayIndex(devNumByVatRate, devNumsByVatRates)
                    netAmmount(ind) = netAmmount(ind) + (Range("P" & i).Value * Range("T" & i).Value)
                    vatAmmount(ind) = vatAmmount(ind) + (Range("P" & i).Value * Range("T" & i).Value * (Range("N" & i).Value / 100))
                End If
            Next i
            
            Set rs = CreateObject("ADODB.Recordset")
            Dim n
            For n = LBound(devNumsByVatRates) To UBound(devNumsByVatRates) - 1
               sqlInsertTaxes = sqlInsertTaxes & queries.insertVatRates( _
                            cnuf, invoiceNum, _
                            CStr(Split(devNumsByVatRates(n), "##")(0)), _
                            CDbl(Split(devNumsByVatRates(n), "##")(1)), _
                            netAmmount(n), _
                            vatAmmount(n), _
                            site, fich, util, n + 1)
            Next n
            
            'Debug.Print sqlInsertTaxes
            rs.Open sqlInsertTaxes, Cn, adOpenStatic
            Set rs = Nothing
            'taxes unosim samo ako imam iznos na stopi poreza
    
    
            'intcfart
            Set rs = CreateObject("ADODB.Recordset")
            sqlInsertLines = ""
            For i = 7 To LastRow - 1
                sqlInsertLines = sqlInsertLines & queries.insertInvoiceLine( _
                            cnuf, invoiceNum, Range("C" & i).Value, Range("E" & i).Value, Range("I" & i).Value, Range("F" & i).Value, Range("N" & i).Value, _
                            CDbl(Range("P" & i).Value), CDbl(Range("T" & i).Value), site, i - 6, fich, util, Range("G" & i).Value)
                'if pn flag true unosimo istu liniju osim što je vat stopa = 0, nc = 0
                'If Range("H" & i).Value = 1 Then
                    'sqlInsertLines = sqlInsertLines & queries.insertInvoiceLine( _
                                'cnuf, invoiceNum, Range("C" & i).Value, Range("E" & i).Value, Range("I" & i).Value, Range("F" & i).Value, 0, _
                                'CDbl(Range("P" & i).Value), CDbl(0), site, i - 6, fich, util, Range("G" & i).Value)
                'End If
            Next i
            
            'Debug.Print sqlInsertLines
            rs.Open sqlInsertLines, Cn, adOpenStatic
            Set rs = Nothing
                           
            insertLog "save_invoice", _
            "{ site: [" & site & "]" _
            & ", sup: [" & sup & "]" _
            & ", invoice: " & invoice & "" _
            & ", util: [" & util & "]" _
            & " }", CStr(sqlInsertLines)
            
            
            Cn.Close
            Set Cn = Nothing
            
            MsgBox "Raèun je uspješno prebaèen u GOLD EDI suèelje!", vbOKOnly, "Upozorenje"
                
            
        ElseIf ans = 7 Then
            'NO
        End If
    Else
        MsgBox "Raèun je potrebno svesti unutar tehnièke tolerance!", vbOKOnly, "Upozorenje"
    End If
    
        
    ActiveSheet.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
    False, AllowSorting:=True, AllowFiltering:=True
    ActiveSheet.EnableSelection = xlNoRestrictions
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    
End Sub


