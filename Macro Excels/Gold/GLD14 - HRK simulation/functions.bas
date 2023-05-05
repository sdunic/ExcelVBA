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

Sub loadMPCSeasonData()

    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    globals.setAllowEventHandling False
    
  
    Sheets(1).Select
    Range("A5:AQ" & utils.getLastRow("A")).ClearContents
    
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
    
    SQLStr = queries.GetMPCdata
    'Debug.Print (SQLStr)
    
    insertLog "load_MPCData", _
    "{ date: " & Date _
    & " }", CStr(SQLStr)
    
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLStr, Cn, adOpenStatic
    
    If rs.EOF = False Then
    Dim row As Long
    row = 5
    Do Until rs.EOF = True
    
        Range("A" & row).Value = rs(0) 'CEXR
        Range("B" & row).Value = rs(1) 'CEXV
        Range("C" & row).Value = rs(2) 'CINV
        Range("D" & row).Value = rs(3) 'NAZIV
        Range("E" & row).Value = rs(4) 'BARKOD
        Range("F" & row).Value = rs(5) 'NIVO_1
        Range("G" & row).Value = rs(6) 'NIVO_1_OPIS
        Range("H" & row).Value = rs(7) 'NIVO_2
        Range("I" & row).Value = rs(8) 'NIVO_2_OPIS
        Range("J" & row).Value = rs(9) 'NIVO_3
        Range("K" & row).Value = rs(10) 'NIVO_3_OPIS
        Range("L" & row).Value = rs(11) 'NIVO_4
        Range("M" & row).Value = rs(12) 'NIVO_4_OPIS
        Range("N" & row).Value = rs(13) 'NIVO_5
        Range("O" & row).Value = rs(14) 'NIVO_5_OPIS
        Range("P" & row).Value = rs(15) 'DATUM
        Range("Q" & row).Value = rs(16) 'OPIS
        Range("R" & row).Value = rs(17) 'SVOSJSTVO
        Range("S" & row).Value = rs(18) 'OSNOVNA_CIJENA
        Range("T" & row).Value = rs(19) 'MPC_A_NTAR
        Range("U" & row).Value = rs(20) 'MPC_A_CIJENA
        Range("W" & row).Value = rs(22) 'MPC_B_NTAR
        Range("X" & row).Value = rs(23) 'MPC_B_CIJENA
        Range("Z" & row).Value = rs(25) 'MPC_C_NTAR
        Range("AA" & row).Value = rs(26) 'MPC_C_CIJENA
        Range("AC" & row).Value = rs(28) 'MPC_D_NTAR
        Range("AD" & row).Value = rs(29) 'MPC_D_CIJENA
        Range("AF" & row).Value = rs(31) 'MPC_S1_NTAR
        Range("AG" & row).Value = rs(32) 'MPC_S1_CIJENA
        Range("AI" & row).Value = rs(34) 'MPC_S2_NTAR
        Range("AJ" & row).Value = rs(35) 'MPC_S2_CIJENA
        Range("AL" & row).Value = rs(37) 'MPC_S3_NTAR
        Range("AM" & row).Value = rs(38) 'MPC_S3_CIJENA
        Range("AO" & row).Value = rs(40) 'MPC_KAMP_NTAR
        Range("AP" & row).Value = rs(41) 'MPC_KAMP_CIJENA
                
        row = row + 1
        rs.MoveNext
    Loop
    Else
        MsgBox "Ne postoje podaci u GOLD-u. Javite se administratoru!", vbOKOnly, "Informacija"
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


Sub initFormulas()

    Range("V" & 5).Formula2R1C1 = "=CalculatePrice(RC[-2],RC17,RC18,RC19,0)" 'PRIJEDLOG_MPC_A_CIJENA
    Range("Y" & 5).Formula2R1C1 = "=CalculatePrice(RC[-2],RC17,RC18,RC22,RC[-3])" 'PRIJEDLOG_MPC_B_CIJENA
    Range("AB" & 5).Formula2R1C1 = "=CalculatePrice(RC[-2],RC17,RC18,RC22,RC[-3])" 'PRIJEDLOG_MPC_C_CIJENA
    Range("AE" & 5).Formula2R1C1 = "=CalculatePrice(RC[-2],RC17,RC18,RC22,RC[-3])" 'PRIJEDLOG_MPC_D_CIJENA
    Range("AH" & 5).Formula2R1C1 = "=CalculatePrice(RC[-2],RC17,RC18,RC22,RC[-3])" 'PRIJEDLOG_MPC_S1_CIJENA
    Range("AK" & 5).Formula2R1C1 = "=CalculatePrice(RC[-2],RC17,RC18,RC22,RC[-3])" 'PRIJEDLOG_MPC_S2_CIJENA
    Range("AN" & 5).Formula2R1C1 = "=CalculatePrice(RC[-2],RC17,RC18,RC22,RC[-3])" 'PRIJEDLOG_MPC_S3_CIJENA
    Range("AQ" & 5).Formula2R1C1 = "=CalculatePrice(RC[-2],RC17,RC18,RC22,RC[-3])" 'PRIJEDLOG_MPC_KAMP_CIJENA
    
    
    Range("AR" & 5).Formula2R1C1 = "=ROUND(RC[-22]/7.5345, 3)" 'PRIJEDLOG_MPC_A_CIJENA_EUR
    Range("AS" & 5).Formula2R1C1 = "=ROUND(RC[-20]/7.5345, 3)" 'PRIJEDLOG_MPC_B_CIJENA_EUR
    Range("AT" & 5).Formula2R1C1 = "=ROUND(RC[-18]/7.5345, 3)" 'PRIJEDLOG_MPC_C_CIJENA_EUR
    Range("AU" & 5).Formula2R1C1 = "=ROUND(RC[-16]/7.5345, 3)" 'PRIJEDLOG_MPC_D_CIJENA_EUR
    Range("AV" & 5).Formula2R1C1 = "=ROUND(RC[-14]/7.5345, 3)" 'PRIJEDLOG_MPC_S1_CIJENA_EUR
    Range("AW" & 5).Formula2R1C1 = "=ROUND(RC[-12]/7.5345, 3)" 'PRIJEDLOG_MPC_S2_CIJENA_EUR
    Range("AX" & 5).Formula2R1C1 = "=ROUND(RC[-10]/7.5345, 3)" 'PRIJEDLOG_MPC_S3_CIJENA_EUR
    Range("AY" & 5).Formula2R1C1 = "=ROUND(RC[-8]/7.5345, 3)" 'PRIJEDLOG_MPC_KAMP_CIJENA_EUR
    
    
End Sub


