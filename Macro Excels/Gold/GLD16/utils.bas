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
    If val = 0 Then
        getDateString = "NULL"
    Else
        getDateString = "to_date(''" & Format(Day(val), "00") & "-" & Format(Month(val), "00") & "-" & Year(val) & "'',''DD-MM-YYYY'')"
    End If
End Function

Function getPriceValue(val As Variant) As String
    If IsEmpty(val) Then
        getPriceValue = "NULL"
    Else
        getPriceValue = Replace(CStr(CDbl(Replace(val, ".", ","))), ",", ".")
    End If
End Function

Sub allowEventHandling()

    globals.setAllowEventHandling (True)
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault

End Sub


Public Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next i
    IsInArray = False

End Function

Public Function GetArrayIndex(stringToBeFound As String, arr As Variant) As Integer
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = stringToBeFound Then
            GetArrayIndex = i
            Exit Function
        End If
    Next i
    GetArrayIndex = -1

End Function


