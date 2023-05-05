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

Sub initDocument()
    initSheet2
    initSheet1
End Sub

Sub initSheet1()
    ActiveWorkbook.Sheets(1).Unprotect
    
    ActiveWorkbook.Sheets(1).Select
    Range("C2").Value = db.getDocVersion
    Range("C5").Value = utils.getUserName
    Range("C6").ClearContents
    Range("C6").Activate
    Range("C7").ClearContents
    Range("C8").ClearContents
    Range("C9").ClearContents
    Range("C10").ClearContents
  
    ActiveWorkbook.Sheets(1).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Sub initSheet2()
    ActiveWorkbook.Sheets(2).Unprotect
    
    ActiveWorkbook.Sheets(2).Select
    Range("C2").Value = db.getDocVersion
    Range("C5").Value = utils.getUserName
    Range("C6").ClearContents
    Range("C6").Activate
    Range("C8").ClearContents
    Range("C9").ClearContents
    Range("C10").ClearContents
    Range("C11").ClearContents
    Range("F8").ClearContents
    Range("F9").ClearContents
    Range("F10").ClearContents
    Range("F11").ClearContents

    ActiveWorkbook.Sheets(2).Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
End Sub

Sub insertR1Client()
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    
    If oibOk(Range("C6").Value) Then
        
        If (Len(Range("C7").Value) > 0 And Len(Range("C8").Value) > 0 And Len(Range("C9").Value) > 0 And Len(Range("C10").Value) > 0) Then
            ans = MsgBox("Jeste li sigurni da želite kreirati R1 kupca u sustavu?", vbYesNo, "Upozorenje")
            If ans = 6 Then
            
            Set Cn = CreateObject("ADODB.Connection")
            Cn.ConnectionTimeout = 1000
            Cn.commandtimeout = 1000
            Cn.Open db.getConnectionString
            
            SQLStr = queries.getR1ClientGold(Range("C6").Value)
            'Debug.Print SQLStr
            Set rs = CreateObject("ADODB.Recordset")
            rs.Open SQLStr, Cn, adOpenStatic
            
            If rs.EOF = True Then
                goldInsertR1Client
                saopInsertR1Client
                MsgBox "R1 kupac je uspješno ubaèen u sustav!", vbOKOnly, "Informacija"
                initSheet1
            Else
                MsgBox "Kupac s OIB brojem " & Range("C6").Value & " veæ postoji u sustavu!", vbOKOnly, "Informacija"
                insertLog "existing_R1_client", _
                "{ oib: " & Range("C6").Value _
                & ", name: " & Range("C7").Value _
                & ", address: " & Range("C8").Value _
                & ", zipCode: " & Range("C9").Value _
                & ", city: " & Range("C10").Value _
                & " }", ""
                Range("C6").Activate
            End If
            
            Set rs = Nothing
            Cn.Close
            Set Cn = Nothing
                
                
            ElseIf ans = 7 Then
                'NO
            End If
        Else
        
            MsgBox "Potrebno je popuniti sva polja!", vbOKOnly, "Informacija"
            If Len(Range("C10").Value) = 0 Then
                Range("C10").Activate
            End If
            
            If Len(Range("C9").Value) = 0 Then
                Range("C9").Activate
            End If
            
            If Len(Range("C8").Value) = 0 Then
                Range("C8").Activate
            End If
            
            If Len(Range("C7").Value) = 0 Then
                Range("C7").Activate
            End If
        
        End If
    
    Else
        MsgBox "Upisan je pogrešan OIB!", vbOKOnly, "Greška"
        Range("C6").Activate
    End If
    

    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
End Sub

Sub goldInsertR1Client()

        'YES
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        

        SQLinsertR1Client = queries.insertR1ClientGold(utils.getUserName, utils.removeCharacters(Range("C6").Value), Range("C7").Value, Range("C8").Value, Range("C9").Value, Range("C10").Value)
        'Debug.Print SQLinsertR1Client
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLinsertR1Client, Cn, adOpenStatic
        Set rs = Nothing
        
        Cn.Close
        Set Cn = Nothing
        
        insertLog "insert_R1_client_GOLD", _
        "{ oib: " & Range("C6").Value _
        & ", name: " & Range("C7").Value _
        & ", address: " & Range("C8").Value _
        & ", zipCode: " & Range("C9").Value _
        & ", city: " & Range("C10").Value _
        & " }", CStr(SQLinsertR1Client)
        
End Sub

Sub saopInsertR1Client()
        'YES
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        

        SQLinsertR1Client = queries.insertR1ClientSAOP(utils.getUserName, utils.removeCharacters(Range("C6").Value), Range("C7").Value, Range("C8").Value, Range("C9").Value, Range("C10").Value)
        'Debug.Print SQLinsertR1Client
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLinsertR1Client, Cn, adOpenStatic
        Set rs = Nothing
        
        Cn.Close
        Set Cn = Nothing
        
        insertLog "insert_R1_client_SAOP", _
        "{ oib: " & Range("C6").Value _
        & ", name: " & Range("C7").Value _
        & ", address: " & Range("C8").Value _
        & ", zipCode: " & Range("C9").Value _
        & ", city: " & Range("C10").Value _
        & " }", CStr(SQLinsertR1Client)
        
End Sub


Sub getR1Client()

    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    SQLStr = queries.getR1ClientGold(utils.removeCharacters(Range("C6").Value))
    'Debug.Print SQLStr
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLStr, Cn, adOpenStatic
    
    If rs.EOF = False Then
        Range("C8").Value = rs(1)
        Range("C9").Value = rs(2)
        Range("C10").Value = rs(4)
        Range("C11").Value = rs(3)
    Else
        MsgBox "Kupac s OIB brojem " & Range("C6").Value & " ne postoji u sustavu!", vbOKOnly, "Informacija"
    End If
    
     insertLog "get_R1_client", _
    "{ oib: " & Range("C6").Value _
    & " }", CStr(SQLStr)
    
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing

End Sub


Sub updateR1Client()

    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    
    If oibOk(Range("C6").Value) Then
        
        If (Len(Range("F8").Value) > 0 And Len(Range("F9").Value) > 0 And Len(Range("F10").Value) > 0 And Len(Range("F11").Value) > 0) Then
            ans = MsgBox("Jeste li sigurni da želite ažurirati R1 kupca u sustavu?", vbYesNo, "Upozorenje")
            If ans = 6 Then
            
                goldUpdateR1Client
                saopUpdateR1Client
                MsgBox "R1 kupac je uspješno ažuriran u sustavu!", vbOKOnly, "Informacija"
                initSheet2

            ElseIf ans = 7 Then
                'NO
            End If
        Else
        
            MsgBox "Potrebno je popuniti sva polja!", vbOKOnly, "Informacija"
            If Len(Range("F11").Value) = 0 Then
                Range("F11").Activate
            End If
            
            If Len(Range("F10").Value) = 0 Then
                Range("F10").Activate
            End If
            
            If Len(Range("F9").Value) = 0 Then
                Range("F9").Activate
            End If
            
            If Len(Range("F8").Value) = 0 Then
                Range("F8").Activate
            End If
        
        End If
    
    Else
        MsgBox "Upisan je pogrešan OIB!", vbOKOnly, "Greška"
        Range("C6").Activate
    End If
    

    Application.ScreenUpdating = True
    Application.Cursor = xlDefault

End Sub

Sub saopUpdateR1Client()
        'YES
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        

        SQLupdateR1Client = queries.updateR1ClientSAOP(utils.getUserName, utils.removeCharacters(Range("C6").Value), Range("F8").Value, Range("F9").Value, Range("F10").Value, Range("F11").Value)
        'Debug.Print SQLupdateR1Client
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLupdateR1Client, Cn, adOpenStatic
        Set rs = Nothing
        
        Cn.Close
        Set Cn = Nothing
        
        insertLog "update_R1_client_SAOP", _
        "{ oib: " & Range("C6").Value _
        & ", name: " & Range("F8").Value _
        & ", address: " & Range("F9").Value _
        & ", zipCode: " & Range("F10").Value _
        & ", city: " & Range("F11").Value _
        & " }", CStr(SQLupdateR1Client)
        
End Sub

Sub goldUpdateR1Client()
        'YES
        Set Cn = CreateObject("ADODB.Connection")
        Cn.ConnectionTimeout = 1000
        Cn.commandtimeout = 1000
        Cn.Open db.getConnectionString
        

        SQLupdateR1Client = queries.updateR1ClientGold(utils.getUserName, utils.removeCharacters(Range("C6").Value), Range("F8").Value, Range("F9").Value, Range("F10").Value, Range("F11").Value)
        'Debug.Print SQLupdateR1Client
        
        Set rs = CreateObject("ADODB.Recordset")
        rs.Open SQLupdateR1Client, Cn, adOpenStatic
        Set rs = Nothing
        
        Cn.Close
        Set Cn = Nothing
        
        insertLog "update_R1_client_GOLD", _
        "{ oib: " & Range("C6").Value _
        & ", name: " & Range("F8").Value _
        & ", address: " & Range("F9").Value _
        & ", zipCode: " & Range("F10").Value _
        & ", city: " & Range("F11").Value _
        & " }", CStr(SQLupdateR1Client)
        
End Sub
