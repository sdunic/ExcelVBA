Attribute VB_Name = "ssh"
Sub procesOrders(sites As String)
    Set WSH = CreateObject("WScript.Shell")
    host = db.getSSHIP
    user = db.getSSHuser
    goCEN = ". /opt/GOLD/ref510/central/Profile ;"
    
    psintDate = Replace(CStr(Format(Date, "dd.MM.yy")), ".", "/")
    
    stores = Split(sites, ",")
    
    'psint = ""
    'For m = 0 To (UBound(stores, 1) - LBound(stores, 1))
        'psint = psint + "psint05p psint05p $USERID " & CStr(psintDate) & " " & stores(m) & " -1 -u" & CStr(Left(utils.getUserName, 12)) & " GB 123 ; "
    'Next m
    
    psint = "psint05p psint05p $USERID " & CStr(psintDate) & " -1 -1 -u" & CStr(Left(utils.getUserName, 12)) & " GB 123 > /dev/null ; "
    
    cmd = "ssh -tt " & user & "@" & host & " << EOF " & goCEN & psint & " EOF"

    Debug.Print cmd
    Set wshOut = WSH.Exec(cmd)

    ''' Capture StdOut '''
    While Not wshOut.StdOut.AtEndOfStream
        sShellOutLine = wshOut.StdOut.ReadLine
        'Debug.Print sShellOutLine
        If sShellOutLine <> "" Then
            sShellOut = sShellOut & sShellOutLine & vbCrLf
        End If
    Wend
    
    ''' Capture StdErr '''
    While Not wshOut.StdErr.AtEndOfStream
        sShellOutLineErr = wshOut.StdErr.ReadLine
        'Debug.Print sShellOutLineErr
        If sShellOutLineErr <> "" Then
            sShellOutErr = sShellOutErr & sShellOutLineErr & vbCrLf
        End If
    Wend
    
    'Debug.Print "StdOut: " & sShellOut
    'Debug.Print "StdErr: " & sShellOutErr

End Sub

