Attribute VB_Name = "utils"
'pass je gold1950

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
    getDateString = Format(Day(val), "00") & "-" & Format(Month(val), "00") & "-" & Year(val)
End Function

Function getDateStringProcedure(val As Date) As String
    getDateStringProcedure = Year(val) & Format(Month(val), "00") & Format(Day(val), "00")
End Function

Function getPriceValue(val As Variant) As Double
    getPriceValue = CDbl(Replace(val, ".", ","))
End Function
