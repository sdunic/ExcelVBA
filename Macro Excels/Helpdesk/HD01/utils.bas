Attribute VB_Name = "utils"
Function getLastRow(column As String) As Long
    Dim sht As Worksheet
    Set sht = ActiveSheet
    getLastRow = sht.Cells(sht.Rows.Count, column).End(xlUp).row + 1
End Function


Function removeCharacters(val As String) As String
    Dim i As Integer
    removeCharacters = ""
    
    For i = 1 To Len(val)
        If IsNumeric(Mid(val, i, 1)) = True Then
            removeCharacters = removeCharacters + Mid(val, i, 1)
        End If
    Next
    
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

Sub PasteasValue()
Attribute PasteasValue.VB_ProcData.VB_Invoke_Func = "v\n14"
    If Application.CutCopyMode = xlCopy Then
         Selection.PasteSpecial xlPasteValues
    Else
        Dim CB As Object
        Set CB = CreateObject("New:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}")
        CB.GetFromClipboard
        Selection.Value = CB.GetText
    End If
End Sub

' Excel Macro funkcija za provjeru kontrolne znamenke OIB-a
Public Function oibOk(oib As String) As Boolean
    If IsNumeric(Left(oib, 1)) = False Then
        oibOk = True
        Exit Function
    End If

    Dim a As Integer
    Dim k As Integer
    If 11 = Len(oib) Then
        a = 10
        For i = 1 To Len(oib) - 1 Step 1
            o = Mid$(oib, i, 1)
            a = a + o
            a = a Mod 10
            If 0 = a Then
                a = 10
            End If
            a = a * 2
            a = a Mod 11
            
            If IsNumeric(o) Then
                Sum = o + Sum
            End If
        Next i
        k = 11 - a
        If 10 = k Then
            k = 0
        End If
        If k = Mid$(oib, i, 11) Then
            oibOk = True
        Else
            oibOk = False
        End If
    Else
        oibOk = False
    End If
End Function
