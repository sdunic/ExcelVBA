Attribute VB_Name = "utils"
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

Function getPriceValue(val As Variant) As Double
    getPriceValue = CDbl(Replace(val, ".", ","))
End Function

Sub setValidation(rng As String)
    range(rng).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="NE,DA"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Selection.FormatConditions.Add Type:=xlTextString, String:="DA", _
        TextOperator:=xlContains
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

Sub clearValidation(rng As String)
    range(rng).Select
    Cells.FormatConditions.Delete
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateInputOnly, AlertStyle:=xlValidAlertStop, Operator _
        :=xlBetween
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub


