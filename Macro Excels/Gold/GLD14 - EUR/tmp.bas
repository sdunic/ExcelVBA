Attribute VB_Name = "tmp"
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

Sub calculateIndex()

Application.Cursor = xlWait
Application.ScreenUpdating = False

Dim i As Long, row As Long
row = Range("A:A").Cells(Range("A:A").Rows.Count, "A").End(xlUp).row
Range("AC4:AH" & row).ClearContents

For i = 4 To row
    If UCase(Range("K" & i).Value) = "TOP" Or UCase(Range("K" & i).Value) = "PL" Or UCase(Range("K" & i).Value) = "/" Then
        If Not Application.WorksheetFunction.IsNA(Range("M" & i).Value) And Len(Range("M" & i).Value) > 0 Then
            If Range("O" & i).Value > 0 Then
                Range("AC" & i).Value = Range("V" & i).Value / Range("O" & i).Value
                Range("AC" & i).NumberFormat = "0.00%"
            End If
            
            If Range("P" & i).Value > 0 Then
                Range("AD" & i).Value = Range("W" & i).Value / Range("P" & i).Value
                Range("AD" & i).NumberFormat = "0.00%"
            End If
            
            If Range("Q" & i).Value > 0 Then
                Range("AE" & i).Value = Range("X" & i).Value / Range("Q" & i).Value
                Range("AE" & i).NumberFormat = "0.00%"
            End If
            
            If Range("R" & i).Value > 0 Then
                Range("AF" & i).Value = Range("Y" & i).Value / Range("R" & i).Value
                Range("AF" & i).NumberFormat = "0.00%"
            End If
            
            If Range("S" & i).Value > 0 Then
                Range("AG" & i).Value = Range("Z" & i).Value / Range("S" & i).Value
                Range("AG" & i).NumberFormat = "0.00%"
            End If
            
            If Range("T" & i).Value > 0 Then
                Range("AH" & i).Value = Range("AA" & i).Value / Range("T" & i).Value
                Range("AH" & i).NumberFormat = "0.00%"
            End If
        End If
    End If
Next i

Application.ScreenUpdating = True
Application.Cursor = xlDefault

End Sub

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

Function PriceA(val As Double) As Double
    PriceA = RoundPriceA(val)
End Function

Function PriceB(val As Double, row As Long) As Double
    Dim svojstva() As String
    svojstva = Split(Range("L" & row).Value, ";")
    
    
    If IsInArray("SLADOLED IMPULS", svojstva) Then
        PriceB = val
   
    ElseIf IsInArray("KOŠARICA", svojstva) Then
        
        PriceB = val
        
    ElseIf IsInArray("TOP 500", svojstva) Then
    
        PriceB = val
        
    Else
    
        If val > 89.99 Then
            PriceB = val
        ElseIf val * 1.03 - val <= 2 Then
            PriceB = RoundPrice(val * 1.03)
        Else
            PriceB = RoundPrice(val + 2)
        End If
        
    End If
End Function

Function PriceC(val As Double, prevVal As Double, row As Long) As Double
    Dim svojstva() As String
    svojstva = Split(Range("L" & row).Value, ";")
    
    If IsInArray("SLADOLED IMPULS", svojstva) Then
        PriceC = val

    ElseIf IsInArray("KOŠARICA", svojstva) Then
        
        PriceC = val
        
    ElseIf IsInArray("TOP 500", svojstva) Then
    
        If val * 1.03 - prevVal <= 2 Then
            PriceC = RoundPrice(val * 1.03)
        Else
            PriceC = RoundPrice(prevVal + 2)
        End If
        
    Else
        If val > 89.99 Then
            PriceC = Application.WorksheetFunction.MRound(val + 5, 5) - 0.01
        ElseIf val * 1.06 - prevVal <= 2 Then
            PriceC = RoundPrice(val * 1.06)
        Else
            PriceC = RoundPrice(prevVal + 2)
        End If
        
    End If
End Function

Function PriceD(val As Double, prevVal As Double, row As Long) As Double
    Dim svojstva() As String
    svojstva = Split(Range("L" & row).Value, ";")
    
    If IsInArray("SLADOLED IMPULS", svojstva) Then
        PriceD = val
        
    ElseIf IsInArray("KOŠARICA", svojstva) Then
        PriceD = val
        
    ElseIf IsInArray("TOP 500", svojstva) Then
        PriceD = prevVal
        
    Else
        If val > 89.99 Then
            PriceD = prevVal
        ElseIf val * 1.09 - prevVal <= 2 Then
            PriceD = RoundPrice(val * 1.09)
        Else
            PriceD = RoundPrice(prevVal + 2)
        End If
        
    End If
End Function

Function PriceSez(val As Double, prevVal As Double, row As Long) As Double
    Dim svojstva() As String
    svojstva = Split(Range("L" & row).Value, ";")
    
    
    If IsInArray("SLADOLED IMPULS", svojstva) Then
        PriceSez = val
        
    ElseIf IsInArray("KOŠARICA", svojstva) And IsInArray("SEZONA", svojstva) Then
        
        PriceSez = val
        
    ElseIf (IsInArray("TOP 500", svojstva) Or IsInArray("KOŠARICA", svojstva)) And Not IsInArray("SEZONA", svojstva) Then
    
        If val * 1.09 - prevVal <= 2 Then
            PriceSez = RoundPrice(val * 1.09)
        Else
            PriceSez = RoundPrice(prevVal + 2)
        End If
        
    Else
    
        If val * 1.14 - prevVal <= 2 Then
            PriceSez = RoundPrice(val * 1.14)
        Else
            PriceSez = RoundPrice(prevVal + 2)
        End If
        
    End If
End Function

Sub setNewPrices()

Application.Cursor = xlWait
Application.ScreenUpdating = False

Dim i As Long, row As Long
row = Range("A:A").Cells(Range("A:A").Rows.Count, "A").End(xlUp).row
Range("V4:AA" & row).ClearContents

For i = 4 To row
    If UCase(Range("K" & i).Value) = "TOP" Or UCase(Range("K" & i).Value) = "PL" Or UCase(Range("K" & i).Value) = "/" Then
        If Not Application.WorksheetFunction.IsNA(Range("M" & i).Value) And Len(Range("M" & i).Value) > 0 Then
            Range("V" & i).Value = PriceA(Range("M" & i).Value)
            Range("W" & i).Value = PriceB(Range("V" & i).Value, i)
            Range("X" & i).Value = PriceC(Range("V" & i).Value, Range("W" & i).Value, i)
            Range("Y" & i).Value = PriceD(Range("V" & i).Value, Range("X" & i).Value, i)
            Range("Z" & i).Value = PriceSez(Range("V" & i).Value, Range("Y" & i).Value, i)
        End If
    End If
    
    If Not Application.WorksheetFunction.IsNA(Range("M" & i).Value) And Len(Range("M" & i).Value) > 0 Then
        Range("AA" & i).Value = Round((Range("M" & i).Value - Range("AL" & i).Value) / (1 + (Range("AK" & i).Value / 100)), 2)
    End If
Next i

Application.ScreenUpdating = True
Application.Cursor = xlDefault

End Sub

Sub setNewPriceAndIndex(tmpRow As Long)

Application.Cursor = xlWait
Application.ScreenUpdating = False

Range("V" & tmpRow & ":AH" & tmpRow).ClearContents

If UCase(Range("K" & tmpRow).Value) = "TOP" Or UCase(Range("K" & tmpRow).Value) = "PL" Or UCase(Range("K" & tmpRow).Value) = "/" Then
    If Not Application.WorksheetFunction.IsNA(Range("M" & tmpRow).Value) And Len(Range("M" & tmpRow).Value) > 0 Then
        Range("V" & tmpRow).Value = PriceA(Range("M" & tmpRow).Value)
        Range("W" & tmpRow).Value = PriceB(Range("V" & tmpRow).Value, tmpRow)
        Range("X" & tmpRow).Value = PriceC(Range("V" & tmpRow).Value, Range("W" & tmpRow).Value, tmpRow)
        Range("Y" & tmpRow).Value = PriceD(Range("V" & tmpRow).Value, Range("X" & tmpRow).Value, tmpRow)
        Range("Z" & tmpRow).Value = PriceSez(Range("V" & tmpRow).Value, Range("Y" & tmpRow).Value, tmpRow)
        
        If Len(Range("O" & tmpRow).Value) > 0 Then
            Range("AC" & tmpRow).Value = Range("V" & tmpRow).Value / Range("O" & tmpRow).Value
            Range("AC" & tmpRow).NumberFormat = "0.00%"
        End If
        
        If Len(Range("P" & tmpRow).Value) > 0 Then
            Range("AD" & tmpRow).Value = Range("W" & tmpRow).Value / Range("P" & tmpRow).Value
            Range("AD" & tmpRow).NumberFormat = "0.00%"
        End If
        
        If Len(Range("Q" & tmpRow).Value) > 0 Then
            Range("AE" & tmpRow).Value = Range("X" & tmpRow).Value / Range("Q" & tmpRow).Value
            Range("AE" & tmpRow).NumberFormat = "0.00%"
        End If
        
        If Len(Range("R" & tmpRow).Value) > 0 Then
            Range("AF" & tmpRow).Value = Range("Y" & tmpRow).Value / Range("R" & tmpRow).Value
            Range("AF" & tmpRow).NumberFormat = "0.00%"
        End If
        
        If Len(Range("S" & tmpRow).Value) > 0 Then
            Range("AG" & tmpRow).Value = Range("Z" & tmpRow).Value / Range("S" & tmpRow).Value
            Range("AG" & tmpRow).NumberFormat = "0.00%"
        End If
        
    End If
End If

If Not Application.WorksheetFunction.IsNA(Range("M" & tmpRow).Value) And Len(Range("M" & tmpRow).Value) > 0 Then

    Range("AA" & tmpRow).Value = Round((Range("M" & tmpRow).Value - Range("AL" & tmpRow).Value) / (1 + (Range("AK" & tmpRow).Value / 100)), 2)
    
    If Len(Range("T" & tmpRow).Value) > 0 Then
        Range("AH" & tmpRow).Value = Range("AA" & tmpRow).Value / Range("T" & tmpRow).Value
        Range("AH" & tmpRow).NumberFormat = "0.00%"
    End If
    
End If

Application.ScreenUpdating = True
Application.Cursor = xlDefault

End Sub

Sub createMPC(name As String)

    Dim change As Boolean
    change = False
    Dim path As String
    path = Application.ActiveWorkbook.path
    Dim row As Long
    row = Range("A:A").Cells(Range("A:A").Rows.Count, "A").End(xlUp).row
    Dim sifreArtikla() As String
    sifreArtikla = Split(toArray(Range("A4:A" & row)), ";")
    Dim opis() As String
    opis = Split(toArray(Range("K4:K" & row)), ";")
    
    Dim currentPrices As Variant
    If name = "A" Then
        currentPrices = Range("O4:O" & row)
    ElseIf name = "B" Then
        currentPrices = Range("P4:P" & row)
    ElseIf name = "C" Then
        currentPrices = Range("Q4:Q" & row)
    ElseIf name = "D" Then
        currentPrices = Range("R4:R" & row)
    ElseIf name = "SEZ" Then
        currentPrices = Range("S4:S" & row)
    ElseIf name = "CS" Then
        currentPrices = Range("T4:T" & row)
    End If
    
    Dim newPrices As Variant
    If name = "A" Then
        newPrices = Range("V4:V" & row)
    ElseIf name = "B" Then
        newPrices = Range("W4:W" & row)
    ElseIf name = "C" Then
        newPrices = Range("X4:X" & row)
    ElseIf name = "D" Then
        newPrices = Range("Y4:Y" & row)
    ElseIf name = "SEZ" Then
        newPrices = Range("Z4:Z" & row)
    ElseIf name = "CS" Then
        newPrices = Range("AA4:AA" & row)
    End If
    
    Dim pricelistDate As String
    pricelistDate = Range("B2").Value

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    Set oFile = fso.CreateTextFile(Application.ActiveWorkbook.path & "\" & name & "_" & pricelistDate & ".txt")

    Dim tmpRow As Long
    tmpRow = 4
    Dim size As Long, i As Long
    size = UBound(sifreArtikla, 1) - LBound(sifreArtikla, 1)
    
    For i = 0 To size
        If Range("A" & tmpRow).EntireRow.Hidden = False And (UCase(opis(i)) = "TOP" Or UCase(opis(i)) = "PL" Or UCase(opis(i)) = "/") And Not IsEmpty(newPrices(i + 1, 1)) Then
            If currentPrices(i + 1, 1) <> newPrices(i + 1, 1) Then
                change = True
                oFile.WriteLine sifreArtikla(i) & ";" & newPrices(i + 1, 1)
            End If
        End If
        tmpRow = tmpRow + 1
    Next i
    
    oFile.Close
    If Not change Then
        SetAttr path & "\" & name & "_" & pricelistDate & ".txt", vbNormal
        Kill path & "\" & name & "_" & pricelistDate & ".txt"
    End If
    
    Set fso = Nothing
    Set oFile = Nothing
    
End Sub


Sub setInitialData()
    
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    Dim Cn As Variant
    Dim Server_Name As String, Database_Name As String, User_ID As String, Password As String, SQLStr As String
    Dim rs As Variant
    Dim row As Long
    
    Set rs = CreateObject("ADODB.Recordset")
    
    Server_Name = "thor2\saop"
    Database_Name = "TommyIT"
    User_ID = "OPZApp"
    Password = "L071nK41707!"
    
    Dim slqDatum As String
                    
    Set Cn = CreateObject("ADODB.Connection")
    Cn.ConnectionTimeout = 1000
    Cn.commandtimeout = 1000
    Cn.Open "Driver={SQL Server};Server=" & Server_Name & ";Database=" & Database_Name & ";Uid=" & User_ID & ";Pwd=" & Password & ";"
    
    sqlDatum = Format(Year(CDate(Range("B2").Value)), "0000") & Format(Month(CDate(Range("B2").Value)), "00") & Format(Day(CDate(Range("B2").Value)), "00")
    SQLStr = "EXEC [TommyIT].[INFO].[ArtikliCjenik] @DatumDohvata  = N'" & sqlDatum & " ', @PeriodMjesec = " & Range("E2").Value
    
    rs.Open SQLStr, Cn, adOpenStatic
    
    row = 4
    Do Until rs.EOF = True
        Range("A" & row).Value = rs(0)
        Range("B" & row).Value = rs(1)
        Range("C" & row).Value = rs(4)
        Range("D" & row).Value = rs(5)
        Range("E" & row).Value = rs(6)
        Range("F" & row).Value = rs(7)
        Range("G" & row).Value = rs(2)
        Range("H" & row).Value = rs(3)
        Range("I" & row).Value = rs(8)
        Range("J" & row).Value = rs(9)
        Range("M" & row).Value = rs(10)
        Range("O" & row).Value = rs(10)
        Range("P" & row).Value = rs(11)
        Range("Q" & row).Value = rs(12)
        Range("R" & row).Value = rs(13)
        Range("S" & row).Value = rs(14)
        Range("T" & row).Value = rs(15)
        Range("AK" & row).Value = rs(16)
        Range("AL" & row).Value = rs(17)
        Range("K" & row).Value = rs(18)
        Range("L" & row).Value = rs(19)
        row = row + 1
        rs.MoveNext
    Loop
    rs.Close
    
    Set rs = Nothing
    Cn.Close
    Set Cn = Nothing
        
    
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    
End Sub
