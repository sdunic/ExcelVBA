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

Function getPriceValue(val As Variant) As Double
    getPriceValue = CDbl(Replace(val, ".", ","))
End Function

Sub setPrice(row As Long, dateCol As String, dateVal As Variant, priceCol As String, priceVal As Variant, newPriceCol As String, indexCol As String, priceId As Variant, changesCol As String)
    If Len(dateVal) > 0 Then
        dateVal = CDate(Replace(dateVal, " 00:00:00.0000000", ""))
        Range(dateCol & row).Value = dateVal  'Datum
    End If
    Range(priceCol & row).Value = getPriceValue(priceVal) 'cijena
    'Range(newPriceCol & row).Value =  'nova cijena
    
    setIndex row, indexCol, priceCol, newPriceCol, changesCol
    
    Range(priceCol & 3).Value = priceId
End Sub

Sub setIndex(row As Long, indexCol As String, priceCol As String, newPriceCol As String, changesCol As String)
    If Range(priceCol & row).Value > 0 And Len(Range(newPriceCol & row).Value) > 0 Then
       Range(indexCol & row).Value = Range(newPriceCol & row).Value / Range(priceCol & row).Value
    Else
        Range(indexCol & row).ClearContents
    End If
    
    If globals.getAllowEventHandling = True Then
        setChangedItem row, priceCol, newPriceCol, indexCol, changesCol
    End If
    
End Sub

Sub setChangedItem(row As Long, priceCol As String, newPriceCol As String, indexCol As String, changesCol As String)
    If Range(priceCol & row).Value > 0 And Len(Range(newPriceCol & row).Value) > 0 Then
       Range(changesCol & row).Value = Range(changesCol & row).Value + 1
    ElseIf Range(priceCol & row).Value = 0 And Len(Range(newPriceCol & row).Value) > 0 Then
        Range(changesCol & row).Value = Range(changesCol & row).Value + 1
    End If
    
    If CDbl(Range(newPriceCol & row).Value) = 0 And globals.getAllowEventHandling = True Then
        Range(changesCol & row).Value = Range(changesCol & row).Value - 1
        If Range(changesCol & row).Value = 0 Then
            Range(changesCol & row).ClearContents
        End If
    End If
End Sub
