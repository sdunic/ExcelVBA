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
