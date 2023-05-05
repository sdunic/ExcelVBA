Attribute VB_Name = "globals"
Dim allowEventHandling As Boolean

Dim rowCount As Long
'total broj linija za insert u interface tablicu

Dim rowNumber As Long
'brojaè linija za insert u interface tablicu

Dim oldConditions, futureConditions As Boolean


Sub setAllowEventHandling(val As Boolean)
    allowEventHandling = val
End Sub

Function getAllowEventHandling()
    getAllowEventHandling = allowEventHandling
End Function

Sub setRowCount(val As Long)
    rowCount = val
End Sub

Function getRowCount() As Long
    getRowCount = rowCount
End Function

Sub addRowNumber()
    rowNumber = rowNumber + 1
End Sub

Function getRowNumber()
    getRowNumber = rowNumber
End Function


Sub setOldConditions(val As String)
    oldConditions = False
    If val = "DA" Then
        oldConditions = True
    End If
End Sub

Sub setFutureConditions(val As String)
    futureConditions = False
    If val = "DA" Then
        futureConditions = True
    End If
End Sub

Function getOldCond() As Boolean
    getOldCond = oldConditions
End Function

Function getFutureCond() As Boolean
    getFutureCond = futureConditions
End Function

