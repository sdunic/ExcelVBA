Attribute VB_Name = "Module1"
Function DaysInMonth(Year As Integer, Month As Integer)
    DaysInMonth = DateSerial(Year, Month + 1, 1) - DateSerial(Year, Month, 1)
End Function


Function DateDiff(fromDate As Date, toDate As Date) As Variant

    getYears = 0
    getMonths = 0
    getDays = 0
    
    If fromDate > toDate Then
        Dim temp As Date
        temp = fromDate
        fromDate = toDate
        toDate = temp
    End If
    
    If Year(toDate) <> Year(fromDate) Then
        getYears = Year(toDate) - Year(fromDate)

        If Month(toDate) < Month(fromDate) Then
            getYears = getYears - 1
        ElseIf Month(toDate) = Month(fromDate) And Day(toDate) < Day(fromDate) Then
            getYears = getYears - 1
        End If
    End If
    
    Dim FromDateMonthDays As Integer
    FromDateMonthDays = DaysInMonth(Year(fromDate), Month(fromDate))
    
    Dim ToDateMonthDays As Integer
    ToDateMonthDays = DaysInMonth(Year(toDate), Month(toDate))
    
    getDays = FromDateMonthDays - Day(fromDate) + 1

    If FromDateMonthDays = getDays Then
        getMonths = getMonths + 1
    End If

    Dim tempMonth As Integer
    tempMonth = Month(toDate) - 1
    getMonths = tempMonth - Month(fromDate)

    If getMonths < 0 And (Month(toDate) < Month(fromDate) Or (Month(toDate) = Month(fromDate) And Day(toDate) < Day(fromDate))) Then
        getMonths = getMonths + 12
    End If

    If ToDateMonthDays = Day(toDate) Then
        getMonths = getMonths + 1
    Else
        getDays = getDays + Day(toDate)
    End If

    If FromDateMonthDays <= getDays Then
        getDays = getDays - FromDateMonthDays
        getMonths = getMonths + 1
    End If

    If getMonths >= 12 Then
        getMonths = getMonths - 12
        getYears = getYears + 1
    End If
    
    DateDiff = getYears & ";" & getMonths & ";" & getDays
   
End Function

