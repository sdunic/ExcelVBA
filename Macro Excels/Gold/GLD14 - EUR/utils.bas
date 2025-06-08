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


Function CalculatePrice(ntar As String, opis As String, svojstvo As String, val As Double, prevVal As Double, basicPrice As Double) As Double
    
    Dim maxDiff As Double
    'If basicPrice < 15 Then
    '    maxDiff = 0.5
    'Else
    '    maxDiff = 0.4
    'End If
    
    maxDiff = 0.5
    
    
    If UCase(opis) = "TOP" Or UCase(opis) = "PL" Or UCase(opis) = "/" Or Len(opis) = 0 Then
        If ntar = "7850" Then
            'MPC A
            CalculatePrice = MPC_ROUNDPRICE(val, "A", opis)
        ElseIf ntar = "7800" Then
            'MPC B
            CalculatePrice = MPC_PRICE_B(val, val, svojstvo, maxDiff, opis)
        ElseIf ntar = "7750" Then
            'MPC C
            CalculatePrice = MPC_PRICE_C(val, prevVal, svojstvo, maxDiff, opis)
        ElseIf ntar = "7700" Then
            'MPC D
            CalculatePrice = MPC_PRICE_D(val, prevVal, svojstvo, maxDiff, opis)
        ElseIf ntar = "7650" Then
            'MPC S1 = C
            CalculatePrice = MPC_PRICE_S1(val, prevVal, svojstvo, maxDiff, opis)
        ElseIf ntar = "7651" Then
            'MPC S2 = D
            CalculatePrice = MPC_PRICE_S2(val, prevVal, svojstvo, maxDiff, opis)
        ElseIf ntar = "7652" Then
            'MPC S3
            CalculatePrice = MPC_PRICE_S3(val, prevVal, svojstvo, maxDiff, opis)
        ElseIf ntar = "7649" Then
            'MPC KAMP
            CalculatePrice = MPC_PRICE_KAMP(val, prevVal, svojstvo, maxDiff, opis)
        End If
        
    ElseIf Len(opis) > 0 Then
        If ntar = "7850" Or ntar = "7800" Then
            CalculatePrice = val
        Else
            CalculatePrice = prevVal
        End If
    Else
        CalculatePrice = 0
    End If

End Function


Sub testFunction()

    Dim price As Double
    Dim svojstvo As String
    
    svojstvo = ""
    svojstvo = "KOSARICA"
    svojstvo = "SEZONA"
    'price = 8.49
    'price = 7.39
    price = 9.29
    
    'price = 0.59
    'price = 2.29
    'price = 4.09
    'price = 7.09
    'price = 2.19
    
    
    Dim a, b, c, d, s1, s2, s3, kamp As Double
    
    a = CalculatePrice("7850", "", svojstvo, price, 0, price)
    b = CalculatePrice("7800", "", svojstvo, CDbl(a), CDbl(a), price)
    c = CalculatePrice("7750", "", svojstvo, CDbl(a), CDbl(b), price)
    d = CalculatePrice("7700", "", svojstvo, CDbl(a), CDbl(c), price)
    
    s1 = CalculatePrice("7650", "", svojstvo, CDbl(a), CDbl(d), price)
    s2 = CalculatePrice("7651", "", svojstvo, CDbl(a), CDbl(s1), price)
    s3 = CalculatePrice("7652", "", svojstvo, CDbl(a), CDbl(s2), price)
    kamp = CalculatePrice("7649", "", svojstvo, CDbl(a), CDbl(s3), price)
    
    
    Debug.Print "INPUT: " & price
    Debug.Print "SVOJSTVTO: " & svojstvo
    Debug.Print "MPC cijene:"
    Debug.Print "   A: " & a
    Debug.Print "   B: " & b
    Debug.Print "   C: " & c
    Debug.Print "   D: " & d
    Debug.Print "   S1: " & s1
    Debug.Print "   S2: " & s2
    Debug.Print "   S3: " & s3
    Debug.Print "   KAMP: " & kamp
    Debug.Print "######"
    
End Sub


Private Function MPC_ROUNDPRICE(val As Double, pricelist As String, svojstvo As String) As Double
    
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
    Dim flagBasicRound As Boolean
    flagBasicRound = False
    
    If IsInArray("KOSARICA", svojstva) Or IsInArray("SEZONA", svojstva) Or IsInArray("TOP500", svojstva) Then
        flagBasicRound = True
    End If
    
    MPC_ROUNDPRICE = val
    
    If pricelist = "A" Then
        MPC_ROUNDPRICE = val
        
    ElseIf flagBasicRound = True Then
        If val = 0 Then
            MPC_ROUNDPRICE = 0
        ElseIf (Application.WorksheetFunction.MRound(val, 0.05) - 0.01) * 100 Mod 10 = 4 Then
            MPC_ROUNDPRICE = Application.WorksheetFunction.MRound(val, 0.05)
        Else
            MPC_ROUNDPRICE = Application.WorksheetFunction.MRound(val, 0.05) - 0.01
        End If
    
    ElseIf val >= 9 Then
        'zaokruživanje na 0.49 i 0.99, na gornju 0.09
        'MPC_ROUNDPRICE = Application.WorksheetFunction.Ceiling(val, 0.5) - 0.01
        
        Dim modValue As Integer
        modValue = Application.WorksheetFunction.Floor(val, 0.01) * 100 Mod 100
        
        'zaokruživanje na najbliži 0.29, 0.49, 0.69 i 0.99
        If modValue >= 14 And modValue < 39 Then
            MPC_ROUNDPRICE = Application.WorksheetFunction.Floor(val, 1) + 0.29
            
        ElseIf modValue >= 39 And modValue < 59 Then
            MPC_ROUNDPRICE = Application.WorksheetFunction.Floor(val, 1) + 0.49
            
        ElseIf modValue >= 59 And modValue < 84 Then
            MPC_ROUNDPRICE = Application.WorksheetFunction.Floor(val, 1) + 0.69
            
        ElseIf modValue >= 84 Then
            MPC_ROUNDPRICE = Application.WorksheetFunction.Floor(val, 1) + 0.99
            
        Else 'sluèaj kad je manje od 14, zaokružujemo na .99, ali sa nižom brojkom
            MPC_ROUNDPRICE = Application.WorksheetFunction.Floor(val, 1) - 1 + 0.99
            
        End If
        
        
    ElseIf val >= 4 Then
        'zaokruživanje na 0.29, 0.49, 0.69 i 0.99, na gornju 0.09
        'If (Application.WorksheetFunction.Floor(val, 0.01) * 100 Mod 100) <= 29 Then
            'MPC_ROUNDPRICE = Application.WorksheetFunction.Floor(val, 1) + 0.29
        'ElseIf (Application.WorksheetFunction.Floor(val, 0.01) * 100 Mod 100) <= 49 Then
            'MPC_ROUNDPRICE = Application.WorksheetFunction.Floor(val, 1) + 0.49
        'ElseIf (Application.WorksheetFunction.Floor(val, 0.01) * 100 Mod 100) <= 69 Then
            'MPC_ROUNDPRICE = Application.WorksheetFunction.Floor(val, 1) + 0.69
        'Else
            'MPC_ROUNDPRICE = Application.WorksheetFunction.Floor(val, 1) + 0.99
        'End If
        
        'zaokruživanje na 0,05 i 0,09
        If val = 0 Then
            MPC_ROUNDPRICE = 0
        ElseIf (Application.WorksheetFunction.MRound(val, 0.05) - 0.01) * 100 Mod 10 = 4 Then
            MPC_ROUNDPRICE = Application.WorksheetFunction.MRound(val, 0.05)
        Else
            MPC_ROUNDPRICE = Application.WorksheetFunction.MRound(val, 0.05) - 0.01
        End If
        
    ElseIf val >= 2 Then
        
        'zaokruživanje na 0.29, 0.49, 0.69 i 0.99, na gornju 0.09
        'If (Application.WorksheetFunction.Floor(val, 0.01) * 100 Mod 100) <= 29 Then
            'MPC_ROUNDPRICE = Application.WorksheetFunction.Floor(val, 1) + 0.29
        'ElseIf (Application.WorksheetFunction.Floor(val, 0.01) * 100 Mod 100) <= 49 Then
            'MPC_ROUNDPRICE = Application.WorksheetFunction.Floor(val, 1) + 0.49
        'ElseIf (Application.WorksheetFunction.Floor(val, 0.01) * 100 Mod 100) <= 69 Then
            'MPC_ROUNDPRICE = Application.WorksheetFunction.Floor(val, 1) + 0.69
        'Else
            'MPC_ROUNDPRICE = Application.WorksheetFunction.Floor(val, 1) + 0.99
        'End If
        
        'zaokruživanje na 0,05 i 0,09
        If val = 0 Then
            MPC_ROUNDPRICE = 0
        ElseIf (Application.WorksheetFunction.MRound(val, 0.05) - 0.01) * 100 Mod 10 = 4 Then
            MPC_ROUNDPRICE = Application.WorksheetFunction.MRound(val, 0.05)
        Else
            MPC_ROUNDPRICE = Application.WorksheetFunction.MRound(val, 0.05) - 0.01
        End If
        
    Else
        'less then 2€
        'zaokruživanje na 0,05 i 0,09
        If val = 0 Then
            MPC_ROUNDPRICE = 0
        ElseIf (Application.WorksheetFunction.MRound(val, 0.05) - 0.01) * 100 Mod 10 = 4 Then
            MPC_ROUNDPRICE = Application.WorksheetFunction.MRound(val, 0.05)
        Else
            MPC_ROUNDPRICE = Application.WorksheetFunction.MRound(val, 0.05) - 0.01
        End If
         
    End If
    
End Function



Private Function MPC_PRICE_B(val As Double, prevVal As Double, svojstvo As String, maxDiff As Double, opis As String) As Double
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
    
    
    If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_B = val
    
    ElseIf IsInArray("KOSARICA", svojstva) And IsInArray("SEZONA", svojstva) Then
        MPC_PRICE_B = val
   
    ElseIf IsInArray("KOSARICA", svojstva) Then
        MPC_PRICE_B = val
        
    ElseIf IsInArray("TOP500", svojstva) Then
        MPC_PRICE_B = val
        
    Else
        If val >= 20 Then
            If val * 1.04 - prevVal <= maxDiff Then
                MPC_PRICE_B = MPC_ROUNDPRICE(val * 1.04, "B", svojstvo)
            Else
                MPC_PRICE_B = MPC_ROUNDPRICE(prevVal + maxDiff, "B", svojstvo)
            End If
        ElseIf val >= 7 Then
            If val * 1.045 - prevVal <= maxDiff Then
                MPC_PRICE_B = MPC_ROUNDPRICE(val * 1.045, "B", svojstvo)
            Else
                MPC_PRICE_B = MPC_ROUNDPRICE(prevVal + maxDiff, "B", svojstvo)
            End If
        ElseIf val >= 5 Then
            If val * 1.055 - prevVal <= maxDiff Then
                MPC_PRICE_B = MPC_ROUNDPRICE(val * 1.055, "B", svojstvo)
            Else
                MPC_PRICE_B = MPC_ROUNDPRICE(prevVal + maxDiff, "B", svojstvo)
            End If
        ElseIf val >= 2 Then
            If val * 1.06 - prevVal <= maxDiff Then
                MPC_PRICE_B = MPC_ROUNDPRICE(val * 1.06, "B", svojstvo)
            Else
                MPC_PRICE_B = MPC_ROUNDPRICE(prevVal + maxDiff, "B", svojstvo)
            End If
        Else 'less then 2€
            If val * 1.075 - val <= maxDiff Then
                MPC_PRICE_B = MPC_ROUNDPRICE(val * 1.075, "B", svojstvo)
            Else
                MPC_PRICE_B = MPC_ROUNDPRICE(prevVal + maxDiff, "B", svojstvo)
            End If
        End If
        
        
    End If
End Function



Private Function MPC_PRICE_C(val As Double, prevVal As Double, svojstvo As String, maxDiff As Double, opis As String) As Double
    'val - nova osnovna cijena za izraèun
    'prevVal - cijena prethodnog cjenika za usporedbu
    
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")

    Dim top500_cijene() As Variant
    top500_cijene() = Array(0.6, 0.5, 0.46, 0.41, 0.4, 0.36, 0.31, 0.3, 0.26, 0.2, 0.16, 0.12, 0.11, 0.1, 0.02, 0.01, 0.06, 0.07, 0.21, 0.7, 0.8)
    Dim woSvojstvo_cijene() As Variant
    woSvojstvo_cijene() = Array(0.38, 0.34, 0.29, 0.2, 0.19, 0.15, 0.11, 0.1, 0.06, 0.02, 0.01)
    
    
   If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_C = val
    
    ElseIf IsInArray("KOSARICA", svojstva) And IsInArray("SEZONA", svojstva) Then
        MPC_PRICE_C = val

    ElseIf IsInArray("KOSARICA", svojstva) Then
        MPC_PRICE_C = val
        
    ElseIf IsInArray("TOP500", svojstva) Then
        'TODO
        If IsInArray(val, top500_cijene) Then
            MPC_PRICE_C = prevVal
        ElseIf val * 1.03 - prevVal <= maxDiff Then
            MPC_PRICE_C = MPC_ROUNDPRICE(val * 1.03, "C", svojstvo)
        Else
            MPC_PRICE_C = MPC_ROUNDPRICE(prevVal + maxDiff, "C", svojstvo)
        End If
        
    Else
        If val >= 20 Then
            If val * 1.06 - prevVal <= maxDiff Then
                MPC_PRICE_C = MPC_ROUNDPRICE(val * 1.06, "C", svojstvo)
            Else
                MPC_PRICE_C = MPC_ROUNDPRICE(prevVal + maxDiff, "C", svojstvo)
            End If
        ElseIf val >= 7 Then
            If val * 1.075 - prevVal <= maxDiff Then
                MPC_PRICE_C = MPC_ROUNDPRICE(val * 1.075, "C", svojstvo)
            Else
                MPC_PRICE_C = MPC_ROUNDPRICE(prevVal + maxDiff, "C", svojstvo)
            End If
        ElseIf val >= 5 Then
            If val * 1.085 - prevVal <= maxDiff Then
                MPC_PRICE_C = MPC_ROUNDPRICE(val * 1.085, "C", svojstvo)
            Else
                MPC_PRICE_C = MPC_ROUNDPRICE(prevVal + maxDiff, "C", svojstvo)
            End If
        ElseIf val >= 2 Then
            If val * 1.1 - prevVal <= maxDiff Then
                MPC_PRICE_C = MPC_ROUNDPRICE(val * 1.1, "C", svojstvo)
            Else
                MPC_PRICE_C = MPC_ROUNDPRICE(prevVal + maxDiff, "C", svojstvo)
            End If
        Else 'less then 2€
            If IsInArray(val, woSvojstvo_cijene) Then
                MPC_PRICE_C = prevVal
            ElseIf val * 1.125 - prevVal <= maxDiff Then
                MPC_PRICE_C = MPC_ROUNDPRICE(val * 1.125, "C", svojstvo)
            Else
                MPC_PRICE_C = MPC_ROUNDPRICE(prevVal + maxDiff, "C", svojstvo)
            End If
        End If
        
    End If
End Function

Private Function MPC_PRICE_D(val As Double, prevVal As Double, svojstvo As String, maxDiff As Double, opis As String) As Double
    'val - nova osnovna cijena za izraèun
    'prevVal - cijena prethodnog cjenika za usporedbu
    
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")

    If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_D = val
    
    ElseIf IsInArray("KOSARICA", svojstva) And IsInArray("SEZONA", svojstva) Then
        MPC_PRICE_D = val
        
    ElseIf IsInArray("KOSARICA", svojstva) Then
        MPC_PRICE_D = val
        
    ElseIf IsInArray("TOP500", svojstva) Then
        MPC_PRICE_D = prevVal
        
    Else
        If val >= 20 Then
            If val * 1.075 - prevVal <= maxDiff Then
                MPC_PRICE_D = MPC_ROUNDPRICE(val * 1.075, "D", svojstvo)
            Else
                MPC_PRICE_D = MPC_ROUNDPRICE(prevVal + maxDiff, "D", svojstvo)
            End If
        ElseIf val >= 7 Then
            If val * 1.09 - prevVal <= maxDiff Then
                MPC_PRICE_D = MPC_ROUNDPRICE(val * 1.09, "D", svojstvo)
            Else
                MPC_PRICE_D = MPC_ROUNDPRICE(prevVal + maxDiff, "D", svojstvo)
            End If
        ElseIf val >= 5 Then
            If val * 1.1 - prevVal <= maxDiff Then
                MPC_PRICE_D = MPC_ROUNDPRICE(val * 1.1, "D", svojstvo)
            Else
                MPC_PRICE_D = MPC_ROUNDPRICE(prevVal + maxDiff, "D", svojstvo)
            End If
        ElseIf val >= 2 Then
            If val * 1.12 - prevVal <= maxDiff Then
                MPC_PRICE_D = MPC_ROUNDPRICE(val * 1.12, "D", svojstvo)
            Else
                MPC_PRICE_D = MPC_ROUNDPRICE(prevVal + maxDiff, "D", svojstvo)
            End If
        Else 'less then 2€
            If val * 1.15 - prevVal <= maxDiff Then
                MPC_PRICE_D = MPC_ROUNDPRICE(val * 1.15, "D", svojstvo)
            Else
                MPC_PRICE_D = MPC_ROUNDPRICE(prevVal + maxDiff, "D", svojstvo)
            End If
        End If
        
    End If
End Function


Private Function MPC_PRICE_S1(val As Double, prevVal As Double, svojstvo As String, maxDiff As Double, opis As String) As Double
    'val - nova osnovna cijena za izraèun
    'prevVal - cijena prethodnog cjenika za usporedbu
    
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
    
    
    If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_S1 = val
        
    ElseIf IsInArray("KOSARICA", svojstva) And IsInArray("SEZONA", svojstva) Then
        MPC_PRICE_S1 = val
        flagBasicRound = True
    
    'iskljuèena pravila prema tasku 1717
    'KOŠARICA i TOP500 se trebaju diferencirati u S1 cjeniku
    'ElseIf IsInArray("KOSARICA", svojstva) Then
        'MPC_PRICE_S1 = val
        
    ElseIf IsInArray("TOP500", svojstva) Then
        If val * 1.05 - prevVal <= maxDiff Then
            MPC_PRICE_S1 = MPC_ROUNDPRICE(val * 1.05, "S1", svojstvo)
        Else
            MPC_PRICE_S1 = MPC_ROUNDPRICE(prevVal + maxDiff, "S1", svojstvo)
        End If
        
    Else
        If val >= 20 Then
            If val * 1.095 - prevVal <= maxDiff Then
                MPC_PRICE_S1 = MPC_ROUNDPRICE(val * 1.095, "S1", svojstvo)
            Else
                MPC_PRICE_S1 = MPC_ROUNDPRICE(prevVal + maxDiff, "S1", svojstvo)
            End If
        ElseIf val >= 7 Then
            If val * 1.11 - prevVal <= maxDiff Then
                MPC_PRICE_S1 = MPC_ROUNDPRICE(val * 1.11, "S1", svojstvo)
            Else
                MPC_PRICE_S1 = MPC_ROUNDPRICE(prevVal + maxDiff, "S1", svojstvo)
            End If
        ElseIf val >= 5 Then
            If val * 1.12 - prevVal <= maxDiff Then
                MPC_PRICE_S1 = MPC_ROUNDPRICE(val * 1.12, "S1", svojstvo)
            Else
                MPC_PRICE_S1 = MPC_ROUNDPRICE(prevVal + maxDiff, "S1", svojstvo)
            End If
        ElseIf val >= 2 Then
            If val * 1.14 - prevVal <= maxDiff Then
                MPC_PRICE_S1 = MPC_ROUNDPRICE(val * 1.14, "S1", svojstvo)
            Else
                MPC_PRICE_S1 = MPC_ROUNDPRICE(prevVal + maxDiff, "S1", svojstvo)
            End If
        Else 'less then 2€
            If val * 1.17 - prevVal <= maxDiff Then
                MPC_PRICE_S1 = MPC_ROUNDPRICE(val * 1.17, "S1", svojstvo)
            Else
                MPC_PRICE_S1 = MPC_ROUNDPRICE(prevVal + maxDiff, "S1", svojstvo)
            End If
        End If
        
    End If
End Function

Private Function MPC_PRICE_S2(val As Double, prevVal As Double, svojstvo As String, maxDiff As Double, opis As String) As Double
    'val - nova osnovna cijena za izraèun
    'prevVal - cijena prethodnog cjenika za usporedbu
    
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
    

    
    If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_S2 = val
        
    ElseIf IsInArray("KOSARICA", svojstva) And IsInArray("SEZONA", svojstva) Then
        MPC_PRICE_S2 = val
        
    'iskljuèena pravila prema tasku 1717
    'KOŠARICA i TOP500 se trebaju diferencirati u S2 cjeniku
    'ElseIf IsInArray("KOSARICA", svojstva) Then
        'MPC_PRICE_S2 = val
        
    ElseIf IsInArray("TOP500", svojstva) Then
        If val * 1.05 - prevVal <= maxDiff Then
            MPC_PRICE_S2 = MPC_ROUNDPRICE(val * 1.05, "S2", svojstvo)
        Else
            MPC_PRICE_S2 = MPC_ROUNDPRICE(prevVal + maxDiff, "S2", svojstvo)
        End If
        
    Else
        If val >= 20 Then
            If val * 1.11 - prevVal <= maxDiff Then
                MPC_PRICE_S2 = MPC_ROUNDPRICE(val * 1.11, "S2", svojstvo)
            Else
                MPC_PRICE_S2 = MPC_ROUNDPRICE(prevVal + maxDiff, "S2", svojstvo)
            End If
        ElseIf val >= 7 Then
            If val * 1.125 - prevVal <= maxDiff Then
                MPC_PRICE_S2 = MPC_ROUNDPRICE(val * 1.125, "S2", svojstvo)
            Else
                MPC_PRICE_S2 = MPC_ROUNDPRICE(prevVal + maxDiff, "S2", svojstvo)
            End If
        ElseIf val >= 5 Then
            If val * 1.14 - prevVal <= maxDiff Then
                MPC_PRICE_S2 = MPC_ROUNDPRICE(val * 1.14, "S2", svojstvo)
            Else
                MPC_PRICE_S2 = MPC_ROUNDPRICE(prevVal + maxDiff, "S2", svojstvo)
            End If
        ElseIf val >= 2 Then
            If val * 1.16 - prevVal <= maxDiff Then
                MPC_PRICE_S2 = MPC_ROUNDPRICE(val * 1.16, "S2", svojstvo)
            Else
                MPC_PRICE_S2 = MPC_ROUNDPRICE(prevVal + maxDiff, "S2", svojstvo)
            End If
        Else 'less then 2€
            If val * 1.19 - prevVal <= maxDiff Then
                MPC_PRICE_S2 = MPC_ROUNDPRICE(val * 1.19, "S2", svojstvo)
            Else
                MPC_PRICE_S2 = MPC_ROUNDPRICE(prevVal + maxDiff, "S2", svojstvo)
            End If
        End If
        
    End If
End Function

Private Function MPC_PRICE_S3(val As Double, prevVal As Double, svojstvo As String, maxDiff As Double, opis As String) As Double
    'val - nova osnovna cijena za izraèun
    'prevVal - cijena prethodnog cjenika za usporedbu
    
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
    

    
    If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_S3 = val
        
    ElseIf IsInArray("KOSARICA", svojstva) And IsInArray("SEZONA", svojstva) Then
        MPC_PRICE_S3 = val
      
    'iskljuèena pravila prema tasku 1717
    'KOŠARICA i TOP500 se trebaju diferencirati u S3 cjeniku
    'ElseIf IsInArray("KOSARICA", svojstva) Then
        'MPC_PRICE_S3 = val
        
    ElseIf IsInArray("TOP500", svojstva) Then
        If val * 1.05 - prevVal <= maxDiff Then
            MPC_PRICE_S3 = MPC_ROUNDPRICE(val * 1.05, "S3", svojstvo)
        Else
            MPC_PRICE_S3 = MPC_ROUNDPRICE(prevVal + maxDiff, "S3", svojstvo)
        End If
        
    Else
        If val >= 20 Then
            If val * 1.125 - prevVal <= maxDiff Then
                MPC_PRICE_S3 = MPC_ROUNDPRICE(val * 1.125, "S3", svojstvo)
            Else
                MPC_PRICE_S3 = MPC_ROUNDPRICE(prevVal + maxDiff, "S3", svojstvo)
            End If
        ElseIf val >= 7 Then
            If val * 1.14 - prevVal <= maxDiff Then
                MPC_PRICE_S3 = MPC_ROUNDPRICE(val * 1.14, "S3", svojstvo)
            Else
                MPC_PRICE_S3 = MPC_ROUNDPRICE(prevVal + maxDiff, "S3", svojstvo)
            End If
        ElseIf val >= 5 Then
            If val * 1.16 - prevVal <= maxDiff Then
                MPC_PRICE_S3 = MPC_ROUNDPRICE(val * 1.16, "S3", svojstvo)
            Else
                MPC_PRICE_S3 = MPC_ROUNDPRICE(prevVal + maxDiff, "S3", svojstvo)
            End If
        ElseIf val >= 2 Then
            If val * 1.18 - prevVal <= maxDiff Then
                MPC_PRICE_S3 = MPC_ROUNDPRICE(val * 1.18, "S3", svojstvo)
            Else
                MPC_PRICE_S3 = MPC_ROUNDPRICE(prevVal + maxDiff, "S3", svojstvo)
            End If
        Else 'less then 2€
            If val * 1.21 - prevVal <= maxDiff Then
                MPC_PRICE_S3 = MPC_ROUNDPRICE(val * 1.21, "S3", svojstvo)
            Else
                MPC_PRICE_S3 = MPC_ROUNDPRICE(prevVal + maxDiff, "S3", svojstvo)
            End If
        End If
        
    End If
End Function

Private Function MPC_PRICE_KAMP(val As Double, prevVal As Double, svojstvo As String, maxDiff As Double, opis As String) As Double
    'val - nova osnovna cijena za izraèun
    'prevVal - cijena prethodnog cjenika za usporedbu
    
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
    
    If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_KAMP = val
        
     ElseIf IsInArray("KOSARICA", svojstva) And IsInArray("SEZONA", svojstva) Then
        MPC_PRICE_KAMP = val
        
    'iskljuèena pravila prema tasku 1717
    'KOŠARICA i TOP500 se trebaju diferencirati u KAMP cjeniku
    'ElseIf IsInArray("KOSARICA", svojstva) Then
        'MPC_PRICE_KAMP = val
        
    ElseIf IsInArray("TOP500", svojstva) Then
        If val * 1.05 - prevVal <= maxDiff Then
            MPC_PRICE_KAMP = MPC_ROUNDPRICE(val * 1.05, "KAMP", svojstvo)
        Else
            MPC_PRICE_KAMP = MPC_ROUNDPRICE(prevVal + maxDiff, "KAMP", svojstvo)
        End If
        
    Else
        If val >= 20 Then
            If val * 1.145 - prevVal <= maxDiff Then
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(val * 1.145, "KAMP", svojstvo)
            Else
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(prevVal + maxDiff, "KAMP", svojstvo)
            End If
        ElseIf val >= 7 Then
            If val * 1.16 - prevVal <= maxDiff Then
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(val * 1.16, "KAMP", svojstvo)
            Else
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(prevVal + maxDiff, "KAMP", svojstvo)
            End If
        ElseIf val >= 5 Then
            If val * 1.18 - prevVal <= maxDiff Then
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(val * 1.18, "KAMP", svojstvo)
            Else
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(prevVal + maxDiff, "KAMP", svojstvo)
            End If
        ElseIf val >= 2 Then
            If val * 1.2 - prevVal <= maxDiff Then
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(val * 1.2, "KAMP", svojstvo)
            Else
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(prevVal + maxDiff, "KAMP", svojstvo)
            End If
        Else 'less then 2€
            If val * 1.23 - prevVal <= maxDiff Then
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(val * 1.23, "KAMP", svojstvo)
            Else
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(prevVal + maxDiff, "KAMP", svojstvo)
            End If
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
