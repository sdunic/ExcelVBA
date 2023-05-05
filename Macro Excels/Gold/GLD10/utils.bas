Attribute VB_Name = "utils"

Sub Protect(sht As Worksheet)
    sht.Protect DrawingObjects:=False, Contents:=True, Scenarios:= _
        False, AllowSorting:=True, AllowFiltering:=True
End Sub
Sub Unprotect(sht As Worksheet)
    sht.Unprotect
End Sub

Function getLastRow(column As String) As Long
    Dim sht As Worksheet
    Set sht = ActiveSheet
    getLastRow = sht.Cells(sht.Rows.Count, column).End(xlUp).row + 1
End Function

Function getUserName() As String
    'domenski korisnik
    getUserName = Environ$("username")
    
    'ime i prezime
    'GetUserName = Application.UserName
End Function

Function getDateString(val As Date) As String
    getDateString = Year(val) & Format(Month(val), "00") & Format(Day(val), "00")
End Function

Function checkNewDocumentVersion() As String
    
    Dim commaSeparator As Boolean

    currentVersion = CStr(getVersionFromDb)
    docVersion = getDocVersion

    If InStr(currentVersion, ",") Then
        currentVersion = Replace(currentVersion, ",", ".")
    End If

    newVersionAvailable = currentVersion <> docVersion
    
    If (newVersionAvailable) Then
        checkNewDocumentVersion = currentVersion
    Else
        checkNewDocumentVersion = ""
    End If
    
    
End Function


Function getVersionFromDb() As Variant

    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open db.getConnectionString
         
    Sqlstr = queries.getDocumentVersion(db.getDocName)
    'Debug.Print (SQLStr)

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open Sqlstr, Cn, adOpenStatic
         
    getVersionFromDb = rs(0)
    
End Function

