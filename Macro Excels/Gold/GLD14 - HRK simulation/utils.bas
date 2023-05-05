Attribute VB_Name = "utils"
Public Function toArray(RNG As Range)
    Dim arr As Variant
    arr = RNG

    With WorksheetFunction
        If UBound(arr, 2) > 1 Then
            toArray = Join((.Index(arr, 1, 0)), ";")
        Else
            toArray = Join(.Transpose(.Index(arr, 0, 1)), ";")
        End If
    End With
End Function

Function IsInArray(ByVal stringToBeFound As String, ByVal arr As Variant) As Boolean
Dim i As Integer, size As Integer
size = UBound(arr, 1) - LBound(arr, 1)
IsInArray = False
For i = 0 To size
    If arr(i) = stringToBeFound Then
        IsInArray = True
    End If
Next i
End Function

Function getLastRow(column As String) As Long
    Dim sht As Worksheet
    Set sht = ActiveSheet
    getLastRow = sht.Cells(sht.Rows.Count, column).End(xlUp).row + 1
End Function

Function getUserName() As String
    getUserName = Environ$("username")
End Function

Function getDateString(val As Date) As String
    getDateString = Format(Day(val), "00") & "-" & Format(Month(val), "00") & "-" & Year(val)
End Function

Function getPriceValue(val As Variant) As Double
    getPriceValue = CDbl(Replace(val, ".", ","))
End Function


Private Function MPC_ROUNDPRICE(val As Double) As Double
    
    MPC_ROUNDPRICE = val
    If val >= 29.99 And val <= 89.99 Then
        MPC_ROUNDPRICE = Round(val, 0) - 0.01
        If Application.WorksheetFunction.RoundDown(val, 0) Mod 10 = 0 Then
            MPC_ROUNDPRICE = MPC_ROUNDPRICE + 1
        End If
    ElseIf val >= 14.99 Then
        If val - Application.WorksheetFunction.RoundDown(val, 0) >= 0.24 And val - Application.WorksheetFunction.RoundDown(val, 0) < 0.74 Then
            MPC_ROUNDPRICE = Application.WorksheetFunction.RoundDown(val, 0) + 0.49
        ElseIf val - Application.WorksheetFunction.RoundDown(val, 0) < 0.24 Then
            MPC_ROUNDPRICE = Application.WorksheetFunction.RoundDown(val, 0) - 0.01
        Else
            MPC_ROUNDPRICE = Application.WorksheetFunction.RoundUp(val, 0) - 0.01
        End If
    Else
        If val - Application.WorksheetFunction.RoundDown(val, 0) >= 0.14 And val - Application.WorksheetFunction.RoundDown(val, 0) < 0.39 Then
            MPC_ROUNDPRICE = Application.WorksheetFunction.RoundDown(val, 0) + 0.29
        ElseIf val - Application.WorksheetFunction.RoundDown(val, 0) >= 0.39 And val - Application.WorksheetFunction.RoundDown(val, 0) < 0.64 Then
            MPC_ROUNDPRICE = Application.WorksheetFunction.RoundDown(val, 0) + 0.49
        ElseIf val - Application.WorksheetFunction.RoundDown(val, 0) >= 0.64 And val - Application.WorksheetFunction.RoundDown(val, 0) < 0.89 Then
            MPC_ROUNDPRICE = Application.WorksheetFunction.RoundDown(val, 0) + 0.79
        ElseIf val - Application.WorksheetFunction.RoundDown(val, 0) > 0.89 Then
            MPC_ROUNDPRICE = Application.WorksheetFunction.RoundUp(val, 0) - 0.01
        Else
            MPC_ROUNDPRICE = Application.WorksheetFunction.RoundDown(val, 0) - 0.01
        End If
    End If
    
End Function


Function CalculatePrice(ntar As String, opis As String, svojstvo As String, val As Double, prevVal As Double) As Double
        
    If UCase(opis) = "TOP" Or UCase(opis) = "PL" Or UCase(opis) = "/" Or Len(opis) = 0 Then
        If ntar = "7850" Then
            'MPC A
            CalculatePrice = MPC_PRICE_A(val)
        ElseIf ntar = "7800" Then
            'MPC B
            CalculatePrice = MPC_PRICE_B(val, svojstvo)
        ElseIf ntar = "7750" Then
            'MPC C
            CalculatePrice = MPC_PRICE_C(val, prevVal, svojstvo)
        ElseIf ntar = "7700" Then
            'MPC D
            CalculatePrice = MPC_PRICE_D(val, prevVal, svojstvo)
        ElseIf ntar = "7650" Then
            'MPC S1 = C
            CalculatePrice = MPC_PRICE_C(val, prevVal, svojstvo)
        ElseIf ntar = "7651" Then
            'MPC S2 = D
            CalculatePrice = MPC_PRICE_D(val, prevVal, svojstvo)
        ElseIf ntar = "7652" Then
            'MPC S3
            CalculatePrice = MPC_PRICE_S3(val, prevVal, svojstvo)
        ElseIf ntar = "7649" Then
            'MPC KAMP
            CalculatePrice = MPC_PRICE_KAMP(val, prevVal, svojstvo)
        
        Else
            CalculatePrice = 0
        End If
        
    End If

End Function


Private Function MPC_PRICE_A(val As Double) As Double
    
    MPC_PRICE_A = val
    If val > 89.99 Then
        MPC_PRICE_A = Application.WorksheetFunction.MRound(val, 5) - 0.01
    ElseIf val >= 29.99 And val <= 89.99 Then
        MPC_PRICE_A = Round(val, 0) - 0.01
        If Application.WorksheetFunction.RoundDown(val, 0) Mod 10 = 0 Then
            MPC_PRICE_A = MPC_PRICE_A + 1
        End If
    ElseIf val >= 14.99 Then
        If val - Application.WorksheetFunction.RoundDown(val, 0) >= 0.24 And val - Application.WorksheetFunction.RoundDown(val, 0) < 0.74 Then
            MPC_PRICE_A = Application.WorksheetFunction.RoundDown(val, 0) + 0.49
        ElseIf val - Application.WorksheetFunction.RoundDown(val, 0) < 0.24 Then
            MPC_PRICE_A = Application.WorksheetFunction.RoundDown(val, 0) - 0.01
        Else
            MPC_PRICE_A = Application.WorksheetFunction.RoundUp(val, 0) - 0.01
        End If
    End If
    
End Function

Private Function MPC_PRICE_B(val As Double, svojstvo As String) As Double
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
    
    
    If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_B = val
   
    ElseIf IsInArray("KOSARICA", svojstva) Then
        MPC_PRICE_B = val
        
    ElseIf IsInArray("TOP500", svojstva) Then
        MPC_PRICE_B = val
        
    Else
        If val > 89.99 Then
            MPC_PRICE_B = val
        ElseIf val * 1.03 - val <= 2 Then
            MPC_PRICE_B = MPC_ROUNDPRICE(val * 1.03)
        Else
            MPC_PRICE_B = MPC_ROUNDPRICE(val + 2)
        End If
        
    End If
End Function


Private Function MPC_PRICE_C(val As Double, prevVal As Double, svojstvo As String) As Double
    'val - nova osnovna cijena za izraèun
    'prevVal - cijena prethodnog cjenika za usporedbu
    
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
    
   If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_C = val

    ElseIf IsInArray("KOSARICA", svojstva) Then
        
        MPC_PRICE_C = val
        
    ElseIf IsInArray("TOP500", svojstva) Then
        If val * 1.03 - prevVal <= 2 Then
            MPC_PRICE_C = MPC_ROUNDPRICE(val * 1.03)
        Else
            MPC_PRICE_C = MPC_ROUNDPRICE(prevVal + 2)
        End If
        
    Else
        If val > 89.99 Then
            MPC_PRICE_C = Application.WorksheetFunction.MRound(val + 5, 5) - 0.01
        ElseIf val * 1.06 - prevVal <= 2 Then
            MPC_PRICE_C = MPC_ROUNDPRICE(val * 1.06)
        Else
            MPC_PRICE_C = MPC_ROUNDPRICE(prevVal + 2)
        End If
        
    End If
End Function

Private Function MPC_PRICE_D(val As Double, prevVal As Double, svojstvo As String) As Double
    'val - nova osnovna cijena za izraèun
    'prevVal - cijena prethodnog cjenika za usporedbu
    
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
    
    If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_D = val
        
    ElseIf IsInArray("KOSARICA", svojstva) Then
        MPC_PRICE_D = val
        
    ElseIf IsInArray("TOP500", svojstva) Then
        MPC_PRICE_D = prevVal
        
    Else
        If val > 89.99 Then
            MPC_PRICE_D = prevVal
        ElseIf val * 1.09 - prevVal <= 2 Then
            MPC_PRICE_D = MPC_ROUNDPRICE(val * 1.09)
        Else
            MPC_PRICE_D = MPC_ROUNDPRICE(prevVal + 2)
        End If
        
    End If
End Function


Private Function MPC_PRICE_S3(val As Double, prevVal As Double, svojstvo As String) As Double
    'val - nova osnovna cijena za izraèun
    'prevVal - cijena prethodnog cjenika za usporedbu
    
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
        
    If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_S3 = val
        
    ElseIf IsInArray("KOSARICA", svojstva) And IsInArray("SEZONA", svojstva) Then
        MPC_PRICE_S3 = val
        
    ElseIf (IsInArray("TOP500", svojstva) Or IsInArray("KOSARICA", svojstva)) And Not IsInArray("SEZONA", svojstva) Then
        If val * 1.09 - prevVal <= 2 Or prevVal = 0 Then
            MPC_PRICE_S3 = MPC_ROUNDPRICE(val * 1.09)
        Else
            MPC_PRICE_S3 = MPC_ROUNDPRICE(prevVal + 2)
        End If
        
    Else
        If val * 1.14 - prevVal <= 2 Then
            MPC_PRICE_S3 = MPC_ROUNDPRICE(val * 1.14)
        Else
            MPC_PRICE_S3 = MPC_ROUNDPRICE(prevVal + 2)
        End If
        
    End If
End Function


Private Function MPC_PRICE_KAMP(val As Double, prevVal As Double, svojstvo As String) As Double
    'val - nova osnovna cijena za izraèun
    'prevVal - cijena prethodnog cjenika za usporedbu
    
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
        
    If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_KAMP = val
        
    ElseIf IsInArray("KOSARICA", svojstva) And IsInArray("SEZONA", svojstva) Then
        MPC_PRICE_KAMP = val
        
    ElseIf (IsInArray("TOP500", svojstva) Or IsInArray("KOSARICA", svojstva)) And Not IsInArray("SEZONA", svojstva) Then
        If val * 1.15 - prevVal <= 2 Or prevVal = 0 Then
            MPC_PRICE_KAMP = MPC_ROUNDPRICE(val * 1.15)
        Else
            MPC_PRICE_KAMP = MPC_ROUNDPRICE(prevVal + 2)
        End If
        
    Else
        If val * 1.2 - prevVal <= 2 Then
            MPC_PRICE_KAMP = MPC_ROUNDPRICE(val * 1.2)
        Else
            MPC_PRICE_KAMP = MPC_ROUNDPRICE(prevVal + 2)
        End If
        
    End If
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
         
    SQLStr = queries.getDocumentVersion(db.getDocName)
    'Debug.Print (SQLStr)

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open SQLStr, Cn, adOpenStatic
         
    getVersionFromDb = rs(0)
    
End Function
