Attribute VB_Name = "queries"
Function getLog(doc_type As String, doc_name As String, doc_version As String, domain_user As String, operation As String, parameters As String, query As String) As String
    getLog = "INSERT INTO [excel].[excel_logovi] (vrsta, naziv, verzija, korisnik, operacija, parametri, datum_vrijeme, sql_upit) VALUES " _
            & "('" & doc_type & "', '" & doc_name & "', '" & doc_version & "', '" & domain_user & "', '" & operation & ""
    
    If Len(parameters) > 0 Then
        getLog = getLog & "', '" & parameters & "',"
    Else
        getLog = getLog & "', " & "NULL" & ","
    End If
    
    getLog = getLog & " current_timestamp"
    
    If Len(query) > 0 Then
        getLog = getLog & ", '" & query & "'"
    Else
        getLog = getLog & ", " & "NULL"
    End If
    
    getLog = getLog & ")"
    
End Function


Function getR1ClientGold(oib As String)

    getR1ClientGold = "EXEC ('"
    getR1ClientGold = getR1ClientGold & "SELECT * FROM TOMMY_R1_CLIENT "
    getR1ClientGold = getR1ClientGold & "WHERE TPOOIB = ''" & UCase(oib) & "''"
    getR1ClientGold = getR1ClientGold & "') at [" + db.getOracleServer + "];"

End Function

Function insertR1ClientGold(user As String, oib As String, name As String, address As String, zipCode As String, city As String)

    insertR1ClientGold = "EXEC ('"
    insertR1ClientGold = insertR1ClientGold & "INSERT INTO TOMMY_R1_CLIENT (TPOOIB, TPOLIBL, TPORUE1, TPOVILL, TPOCODE, TPOUTIL, TPODCRE, TPODMAJ) VALUES "
    insertR1ClientGold = insertR1ClientGold & "(''" & Trim(UCase(oib)) & "'', ''" & Trim(UCase(name)) & "'', ''" & Trim(UCase(address)) & "'', ''" & Trim(UCase(city)) & "'', ''" & Trim(UCase(zipCode)) & "'', ''" & user & "'', sysdate, sysdate)"
    insertR1ClientGold = insertR1ClientGold & "') at [" + db.getOracleServer + "];"

End Function

Function insertR1ClientSAOP(user As String, oib As String, name As String, address As String, zipCode As String, city As String)

    insertR1ClientSAOP = "INSERT INTO THOR2.TommyIT.NCR.LokalneR1Stranke (OIB, Naziv, Adresa, Grad, Unio, Uneseno) VALUES  "
    insertR1ClientSAOP = insertR1ClientSAOP & "('" & Trim(UCase(oib)) & "', '" & Trim(UCase(name)) & "', '" & Trim(UCase(address)) & "', '" & Trim(UCase(zipCode)) & " " & Trim(UCase(city)) & "', '" & user & "', current_timestamp)"

End Function


Function updateR1ClientSAOP(user As String, oib As String, name As String, address As String, zipCode As String, city As String)

    updateR1ClientSAOP = "UPDATE THOR2.TommyIT.NCR.LokalneR1Stranke "
    updateR1ClientSAOP = updateR1ClientSAOP & "SET Naziv = '" & Trim(UCase(name)) & "', Adresa = '" & Trim(UCase(address)) & "', Grad = '" & Trim(UCase(zipCode)) & " " & Trim(UCase(city)) & "', Unio = '" & user & "', Uneseno = current_timestamp"
    updateR1ClientSAOP = updateR1ClientSAOP & " WHERE OIB = '" & UCase(oib) & "'"

End Function

Function updateR1ClientGold(user As String, oib As String, name As String, address As String, zipCode As String, city As String)

    updateR1ClientGold = "EXEC ('"
    updateR1ClientGold = updateR1ClientGold & "UPDATE TOMMY_R1_CLIENT"
    updateR1ClientGold = updateR1ClientGold & " SET TPOLIBL = ''" & Trim(UCase(name)) & "'', TPORUE1 = ''" & Trim(UCase(address)) & "'', TPOVILL = ''" & Trim(UCase(city)) & "'', TPOCODE = ''" & Trim(UCase(zipCode)) & "'', TPOUTIL =''" & user & "'', TPODMAJ = sysdate"
    updateR1ClientGold = updateR1ClientGold & " WHERE TPOOIB = ''" & UCase(oib) & "''"
    updateR1ClientGold = updateR1ClientGold & "') at [" + db.getOracleServer + "];"

End Function
