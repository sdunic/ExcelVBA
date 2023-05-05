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


Function CalculatePrice(ntar As String, opis As String, svojstvo As String, val As Double, prevVal As Double) As Double
    
    Dim maxDiff As Double
    maxDiff = 0.5
    
    If UCase(opis) = "TOP" Or UCase(opis) = "PL" Or UCase(opis) = "/" Or Len(opis) = 0 Then
        If ntar = "7850" Then
            'MPC A
            CalculatePrice = MPC_ROUNDPRICE(val, "A")
        ElseIf ntar = "7800" Then
            'MPC B
            CalculatePrice = MPC_PRICE_B(val, val, svojstvo, maxDiff)
        ElseIf ntar = "7750" Then
            'MPC C
            CalculatePrice = MPC_PRICE_C(val, prevVal, svojstvo, maxDiff)
        ElseIf ntar = "7700" Then
            'MPC D
            CalculatePrice = MPC_PRICE_D(val, prevVal, svojstvo, maxDiff)
        ElseIf ntar = "7650" Then
            'MPC S1 = C
            CalculatePrice = MPC_PRICE_S1(val, prevVal, svojstvo, maxDiff)
        ElseIf ntar = "7651" Then
            'MPC S2 = D
            CalculatePrice = MPC_PRICE_S2(val, prevVal, svojstvo, maxDiff)
        ElseIf ntar = "7652" Then
            'MPC S3
            CalculatePrice = MPC_PRICE_S3(val, prevVal, svojstvo, maxDiff)
        ElseIf ntar = "7649" Then
            'MPC KAMP
            CalculatePrice = MPC_PRICE_KAMP(val, prevVal, svojstvo, maxDiff)
        
        Else
            CalculatePrice = 0
        End If
        
    End If

End Function


Sub testFunction()

    Dim price As Double
    
    'price = 0.59
    'price = 2.29
    'price = 4.09
    'price = 7.09
    price = 20.09
    
    Dim a, b, c, d, s1, s2, s3, kamp As Double
    
    a = CalculatePrice("7850", "", "", price, 0)
    b = CalculatePrice("7800", "", "", CDbl(a), CDbl(a))
    c = CalculatePrice("7750", "", "", CDbl(a), CDbl(b))
    d = CalculatePrice("7700", "", "", CDbl(a), CDbl(c))
    
    s1 = CalculatePrice("7650", "", "", CDbl(a), CDbl(d))
    s2 = CalculatePrice("7651", "", "", CDbl(a), CDbl(s1))
    s3 = CalculatePrice("7652", "", "", CDbl(a), CDbl(s2))
    kamp = CalculatePrice("7649", "", "", CDbl(a), CDbl(s3))
    
    
    Debug.Print "INPUT: " & price
    Debug.Print "A: " & a
    Debug.Print "B: " & b
    Debug.Print "C: " & c
    Debug.Print "D: " & d
    Debug.Print "S1: " & s1
    Debug.Print "S2: " & s2
    Debug.Print "S3: " & s3
    Debug.Print "KAMP: " & kamp
    
    Debug.Print "######"
    
End Sub


Private Function MPC_ROUNDPRICE(val As Double, pricelist As String) As Double
    
    MPC_ROUNDPRICE = val
    
    If val >= 20 Then
        'A,B,C,D,S1,S2,S3,KAMP ROUND UP, zaokruživanje na 0,49 i 0,99
        If pricelist = "A" Or pricelist = "B" Or pricelist = "C" Or pricelist = "D" Or pricelist = "S1" Or pricelist = "S2" Or pricelist = "S3" Or pricelist = "KAMP" Then
            MPC_ROUNDPRICE = Application.WorksheetFunction.Ceiling(val, 0.5) - 0.01
        End If
        
    ElseIf val >= 7 Then
        'A,B,C,D,S1,S2,S3,KAMP ROUND UP, zaokruživanje na 0,09
        If pricelist = "A" Or pricelist = "B" Or pricelist = "C" Or pricelist = "D" Or pricelist = "S1" Or pricelist = "S2" Or pricelist = "S3" Or pricelist = "KAMP" Then
            MPC_ROUNDPRICE = Application.WorksheetFunction.Ceiling(val, 0.1) - 0.01
        End If
        
    ElseIf val >= 4 Then
        'A,B ROUND; C,D,S1,S2,S3,KAMP ROUND UP, zaokruživanje na 0,05 i 0,09
        If pricelist = "A" Or pricelist = "B" Then
            If (Application.WorksheetFunction.MRound(val, 0.05) - 0.01) * 100 Mod 10 = 4 Then
                MPC_ROUNDPRICE = Application.WorksheetFunction.MRound(val, 0.05)
            Else
                MPC_ROUNDPRICE = Application.WorksheetFunction.MRound(val, 0.05) - 0.01
            End If
       ElseIf pricelist = "C" Or pricelist = "D" Or pricelist = "S1" Or pricelist = "S2" Or pricelist = "S3" Or pricelist = "KAMP" Then
            If (Application.WorksheetFunction.Ceiling(val, 0.05) - 0.01) * 100 Mod 10 = 4 Then
                MPC_ROUNDPRICE = Application.WorksheetFunction.Ceiling(val, 0.05)
            Else
                MPC_ROUNDPRICE = Application.WorksheetFunction.Ceiling(val, 0.05) - 0.01
            End If
        End If
        
    ElseIf val >= 2 Then
        'A,B ROUND; C,D,S1,S2,S3,KAMP ROUND UP, zaokruživanje na 0,05 i 0,09
        If pricelist = "A" Or pricelist = "B" Then
            If (Application.WorksheetFunction.MRound(val, 0.05) - 0.01) * 100 Mod 10 = 4 Then
                MPC_ROUNDPRICE = Application.WorksheetFunction.MRound(val, 0.05)
            Else
                MPC_ROUNDPRICE = Application.WorksheetFunction.MRound(val, 0.05) - 0.01
            End If
       ElseIf pricelist = "C" Or pricelist = "D" Or pricelist = "S1" Or pricelist = "S2" Or pricelist = "S3" Or pricelist = "KAMP" Then
            If (Application.WorksheetFunction.Ceiling(val, 0.05) - 0.01) * 100 Mod 10 = 4 Then
                MPC_ROUNDPRICE = Application.WorksheetFunction.Ceiling(val, 0.05)
            Else
                MPC_ROUNDPRICE = Application.WorksheetFunction.Ceiling(val, 0.05) - 0.01
            End If
        End If
        
    Else
        'less then 2€
        'A,B NO ROUND
        'C ROUND, D,S1,S2,S3,KAMP ROUNDUP, zaokruživanje na 0,05 i 0,09
        
        If pricelist = "A" Or pricelist = "B" Then
            MPC_ROUNDPRICE = Round(val, 2)
        ElseIf pricelist = "C" Then
           If (Application.WorksheetFunction.MRound(val, 0.05) - 0.01) * 100 Mod 10 = 4 Then
                MPC_ROUNDPRICE = Application.WorksheetFunction.MRound(val, 0.05)
            Else
                MPC_ROUNDPRICE = Application.WorksheetFunction.MRound(val, 0.05) - 0.01
            End If
         ElseIf pricelist = "D" Or pricelist = "S1" Or pricelist = "S2" Or pricelist = "S3" Or pricelist = "KAMP" Then
            If (Application.WorksheetFunction.Ceiling(val, 0.05) - 0.01) * 100 Mod 10 = 4 Then
                MPC_ROUNDPRICE = Application.WorksheetFunction.Ceiling(val, 0.05)
            Else
                MPC_ROUNDPRICE = Application.WorksheetFunction.Ceiling(val, 0.05) - 0.01
            End If
        End If
       
    End If
    
End Function

Private Function MPC_PRICE_B(val As Double, prevVal As Double, svojstvo As String, maxDiff As Double) As Double
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
    
    
    If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_B = val
   
    ElseIf IsInArray("KOSARICA", svojstva) Then
        MPC_PRICE_B = val
        
    ElseIf IsInArray("TOP500", svojstva) Then
        MPC_PRICE_B = val
        
    Else
        If val >= 20 Then
            If val * 1.025 - prevVal <= maxDiff Then
                MPC_PRICE_B = MPC_ROUNDPRICE(val * 1.025, "B")
            Else
                MPC_PRICE_B = MPC_ROUNDPRICE(prevVal + maxDiff, "B")
            End If
        ElseIf val >= 7 Then
            If val * 1.03 - prevVal <= maxDiff Then
                MPC_PRICE_B = MPC_ROUNDPRICE(val * 1.03, "B")
            Else
                MPC_PRICE_B = MPC_ROUNDPRICE(prevVal + maxDiff, "B")
            End If
        ElseIf val >= 4 Then
            If val * 1.035 - prevVal <= maxDiff Then
                MPC_PRICE_B = MPC_ROUNDPRICE(val * 1.035, "B")
            Else
                MPC_PRICE_B = MPC_ROUNDPRICE(prevVal + maxDiff, "B")
            End If
        ElseIf val >= 2 Then
            If val * 1.04 - prevVal <= maxDiff Then
                MPC_PRICE_B = MPC_ROUNDPRICE(val * 1.04, "B")
            Else
                MPC_PRICE_B = MPC_ROUNDPRICE(prevVal + maxDiff, "B")
            End If
        Else 'less then 2€
            If val * 1.05 - val <= maxDiff Then
                MPC_PRICE_B = MPC_ROUNDPRICE(val * 1.05, "B")
            Else
                MPC_PRICE_B = MPC_ROUNDPRICE(prevVal + maxDiff, "B")
            End If
        End If
        
        
    End If
End Function


Private Function MPC_PRICE_C(val As Double, prevVal As Double, svojstvo As String, maxDiff As Double) As Double
    'val - nova osnovna cijena za izraèun
    'prevVal - cijena prethodnog cjenika za usporedbu
    
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
    
   If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_C = val

    ElseIf IsInArray("KOSARICA", svojstva) Then
        
        MPC_PRICE_C = val
        
    ElseIf IsInArray("TOP500", svojstva) Then
        'TODO
        If val * 1.03 - prevVal <= maxDiff Then
            MPC_PRICE_C = MPC_ROUNDPRICE(val * 1.03, "C")
        Else
            MPC_PRICE_C = MPC_ROUNDPRICE(prevVal + maxDiff, "C")
        End If
        
    Else
        If val >= 20 Then
            If val * 1.05 - prevVal <= maxDiff Then
                MPC_PRICE_C = MPC_ROUNDPRICE(val * 1.05, "C")
            Else
                MPC_PRICE_C = MPC_ROUNDPRICE(prevVal + maxDiff, "C")
            End If
        ElseIf val >= 7 Then
            If val * 1.06 - prevVal <= maxDiff Then
                MPC_PRICE_C = MPC_ROUNDPRICE(val * 1.06, "C")
            Else
                MPC_PRICE_C = MPC_ROUNDPRICE(prevVal + maxDiff, "C")
            End If
        ElseIf val >= 4 Then
            If val * 1.07 - prevVal <= maxDiff Then
                MPC_PRICE_C = MPC_ROUNDPRICE(val * 1.07, "C")
            Else
                MPC_PRICE_C = MPC_ROUNDPRICE(prevVal + maxDiff, "C")
            End If
        ElseIf val >= 2 Then
            If val * 1.08 - prevVal <= maxDiff Then
                MPC_PRICE_C = MPC_ROUNDPRICE(val * 1.08, "C")
            Else
                MPC_PRICE_C = MPC_ROUNDPRICE(prevVal + maxDiff, "C")
            End If
        Else 'less then 2€
            If val * 1.1 - prevVal <= maxDiff Then
                MPC_PRICE_C = MPC_ROUNDPRICE(val * 1.1, "C")
            Else
                MPC_PRICE_C = MPC_ROUNDPRICE(prevVal + maxDiff, "C")
            End If
        End If
        
    End If
End Function

Private Function MPC_PRICE_D(val As Double, prevVal As Double, svojstvo As String, maxDiff As Double) As Double
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
        If val >= 20 Then
            If val * 1.075 - prevVal <= maxDiff Then
                MPC_PRICE_D = MPC_ROUNDPRICE(val * 1.075, "D")
            Else
                MPC_PRICE_D = MPC_ROUNDPRICE(prevVal + maxDiff, "D")
            End If
        ElseIf val >= 7 Then
            If val * 1.09 - prevVal <= maxDiff Then
                MPC_PRICE_D = MPC_ROUNDPRICE(val * 1.09, "D")
            Else
                MPC_PRICE_D = MPC_ROUNDPRICE(prevVal + maxDiff, "D")
            End If
        ElseIf val >= 4 Then
            If val * 1.1 - prevVal <= maxDiff Then
                MPC_PRICE_D = MPC_ROUNDPRICE(val * 1.1, "D")
            Else
                MPC_PRICE_D = MPC_ROUNDPRICE(prevVal + maxDiff, "D")
            End If
        ElseIf val >= 2 Then
            If val * 1.12 - prevVal <= maxDiff Then
                MPC_PRICE_D = MPC_ROUNDPRICE(val * 1.12, "D")
            Else
                MPC_PRICE_D = MPC_ROUNDPRICE(prevVal + maxDiff, "D")
            End If
        Else 'less then 2€
            If val * 1.15 - prevVal <= maxDiff Then
                MPC_PRICE_D = MPC_ROUNDPRICE(val * 1.15, "D")
            Else
                MPC_PRICE_D = MPC_ROUNDPRICE(prevVal + maxDiff, "D")
            End If
        End If
        
    End If
End Function


Private Function MPC_PRICE_S1(val As Double, prevVal As Double, svojstvo As String, maxDiff As Double) As Double
    'val - nova osnovna cijena za izraèun
    'prevVal - cijena prethodnog cjenika za usporedbu
    
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
    
    If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_S1 = val
        
    ElseIf IsInArray("KOSARICA", svojstva) Then
        MPC_PRICE_S1 = val
        
    ElseIf IsInArray("TOP500", svojstva) Then
        MPC_PRICE_S1 = prevVal
        
    Else
        If val >= 20 Then
            If val * 1.095 - prevVal <= maxDiff Then
                MPC_PRICE_S1 = MPC_ROUNDPRICE(val * 1.095, "S1")
            Else
                MPC_PRICE_S1 = MPC_ROUNDPRICE(prevVal + maxDiff, "S1")
            End If
        ElseIf val >= 7 Then
            If val * 1.11 - prevVal <= maxDiff Then
                MPC_PRICE_S1 = MPC_ROUNDPRICE(val * 1.11, "S1")
            Else
                MPC_PRICE_S1 = MPC_ROUNDPRICE(prevVal + maxDiff, "S1")
            End If
        ElseIf val >= 4 Then
            If val * 1.12 - prevVal <= maxDiff Then
                MPC_PRICE_S1 = MPC_ROUNDPRICE(val * 1.12, "S1")
            Else
                MPC_PRICE_S1 = MPC_ROUNDPRICE(prevVal + maxDiff, "S1")
            End If
        ElseIf val >= 2 Then
            If val * 1.14 - prevVal <= maxDiff Then
                MPC_PRICE_S1 = MPC_ROUNDPRICE(val * 1.14, "S1")
            Else
                MPC_PRICE_S1 = MPC_ROUNDPRICE(prevVal + maxDiff, "S1")
            End If
        Else 'less then 2€
            If val * 1.17 - prevVal <= maxDiff Then
                MPC_PRICE_S1 = MPC_ROUNDPRICE(val * 1.17, "S1")
            Else
                MPC_PRICE_S1 = MPC_ROUNDPRICE(prevVal + maxDiff, "S1")
            End If
        End If
        
    End If
End Function

Private Function MPC_PRICE_S2(val As Double, prevVal As Double, svojstvo As String, maxDiff As Double) As Double
    'val - nova osnovna cijena za izraèun
    'prevVal - cijena prethodnog cjenika za usporedbu
    
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
    
    If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_S2 = val
        
    ElseIf IsInArray("KOSARICA", svojstva) Then
        MPC_PRICE_S2 = val
        
    ElseIf IsInArray("TOP500", svojstva) Then
        MPC_PRICE_S2 = prevVal
        
    Else
        If val >= 20 Then
            If val * 1.11 - prevVal <= maxDiff Then
                MPC_PRICE_S2 = MPC_ROUNDPRICE(val * 1.11, "S2")
            Else
                MPC_PRICE_S2 = MPC_ROUNDPRICE(prevVal + maxDiff, "S2")
            End If
        ElseIf val >= 7 Then
            If val * 1.125 - prevVal <= maxDiff Then
                MPC_PRICE_S2 = MPC_ROUNDPRICE(val * 1.125, "S2")
            Else
                MPC_PRICE_S2 = MPC_ROUNDPRICE(prevVal + maxDiff, "S2")
            End If
        ElseIf val >= 4 Then
            If val * 1.14 - prevVal <= maxDiff Then
                MPC_PRICE_S2 = MPC_ROUNDPRICE(val * 1.14, "S2")
            Else
                MPC_PRICE_S2 = MPC_ROUNDPRICE(prevVal + maxDiff, "S2")
            End If
        ElseIf val >= 2 Then
            If val * 1.16 - prevVal <= maxDiff Then
                MPC_PRICE_S2 = MPC_ROUNDPRICE(val * 1.16, "S2")
            Else
                MPC_PRICE_S2 = MPC_ROUNDPRICE(prevVal + maxDiff, "S2")
            End If
        Else 'less then 2€
            If val * 1.19 - prevVal <= maxDiff Then
                MPC_PRICE_S2 = MPC_ROUNDPRICE(val * 1.19, "S2")
            Else
                MPC_PRICE_S2 = MPC_ROUNDPRICE(prevVal + maxDiff, "S2")
            End If
        End If
        
    End If
End Function

Private Function MPC_PRICE_S3(val As Double, prevVal As Double, svojstvo As String, maxDiff As Double) As Double
    'val - nova osnovna cijena za izraèun
    'prevVal - cijena prethodnog cjenika za usporedbu
    
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
    
    If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_S3 = val
        
    ElseIf IsInArray("KOSARICA", svojstva) Then
        MPC_PRICE_S3 = val
        
    ElseIf IsInArray("TOP500", svojstva) Then
        MPC_PRICE_S3 = prevVal
        
    Else
        If val >= 20 Then
            If val * 1.125 - prevVal <= maxDiff Then
                MPC_PRICE_S3 = MPC_ROUNDPRICE(val * 1.125, "S3")
            Else
                MPC_PRICE_S3 = MPC_ROUNDPRICE(prevVal + maxDiff, "S3")
            End If
        ElseIf val >= 7 Then
            If val * 1.14 - prevVal <= maxDiff Then
                MPC_PRICE_S3 = MPC_ROUNDPRICE(val * 1.14, "S3")
            Else
                MPC_PRICE_S3 = MPC_ROUNDPRICE(prevVal + maxDiff, "S3")
            End If
        ElseIf val >= 4 Then
            If val * 1.16 - prevVal <= maxDiff Then
                MPC_PRICE_S3 = MPC_ROUNDPRICE(val * 1.16, "S3")
            Else
                MPC_PRICE_S3 = MPC_ROUNDPRICE(prevVal + maxDiff, "S3")
            End If
        ElseIf val >= 2 Then
            If val * 1.18 - prevVal <= maxDiff Then
                MPC_PRICE_S3 = MPC_ROUNDPRICE(val * 1.18, "S3")
            Else
                MPC_PRICE_S3 = MPC_ROUNDPRICE(prevVal + maxDiff, "S3")
            End If
        Else 'less then 2€
            If val * 1.21 - prevVal <= maxDiff Then
                MPC_PRICE_S3 = MPC_ROUNDPRICE(val * 1.21, "S3")
            Else
                MPC_PRICE_S3 = MPC_ROUNDPRICE(prevVal + maxDiff, "S3")
            End If
        End If
        
    End If
End Function

Private Function MPC_PRICE_KAMP(val As Double, prevVal As Double, svojstvo As String, maxDiff As Double) As Double
    'val - nova osnovna cijena za izraèun
    'prevVal - cijena prethodnog cjenika za usporedbu
    
    Dim svojstva() As String
    svojstva = Split(svojstvo, ";")
    
    If IsInArray("IMPULS", svojstva) And IsInArray("SLADOLED", svojstva) Then
        MPC_PRICE_KAMP = val
        
    ElseIf IsInArray("KOSARICA", svojstva) Then
        MPC_PRICE_KAMP = val
        
    ElseIf IsInArray("TOP500", svojstva) Then
        MPC_PRICE_KAMP = prevVal
        
    Else
        If val >= 20 Then
            If val * 1.145 - prevVal <= maxDiff Then
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(val * 1.145, "KAMP")
            Else
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(prevVal + maxDiff, "KAMP")
            End If
        ElseIf val >= 7 Then
            If val * 1.16 - prevVal <= maxDiff Then
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(val * 1.16, "KAMP")
            Else
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(prevVal + maxDiff, "KAMP")
            End If
        ElseIf val >= 4 Then
            If val * 1.18 - prevVal <= maxDiff Then
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(val * 1.18, "KAMP")
            Else
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(prevVal + maxDiff, "KAMP")
            End If
        ElseIf val >= 2 Then
            If val * 1.2 - prevVal <= maxDiff Then
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(val * 1.2, "KAMP")
            Else
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(prevVal + maxDiff, "KAMP")
            End If
        Else 'less then 2€
            If val * 1.23 - prevVal <= maxDiff Then
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(val * 1.23, "KAMP")
            Else
                MPC_PRICE_KAMP = MPC_ROUNDPRICE(prevVal + maxDiff, "KAMP")
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
