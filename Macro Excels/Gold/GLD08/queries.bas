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

Function searchMSNodes(code As String, name As String) As String
    searchMSNodes = "EXEC ('" _
            & "SELECT DISTINCT pkstrucobj.get_cext(123, objcint) MS_CODE, pkstrucobj.get_desc(123, objcint, ''HR'') MS_DESC " _
            & "FROM strucrel " _
            & "WHERE NOT EXISTS (SELECT 1 FROM artrac WHERE artcinr = objcint) "
    
    If Len(code) > 0 Then
        searchMSNodes = searchMSNodes & "AND pkstrucobj.get_cext(123, objcint) LIKE ''" & UCase(code) & "'' "
    End If
    
    If Len(name) > 0 Then
        searchMSNodes = searchMSNodes & "AND pkstrucobj.get_desc(123, objcint, ''HR'') LIKE ''" & UCase(name) & "'' "
    End If
            
    searchMSNodes = searchMSNodes & "ORDER BY 1') at [" + db.getOracleServer + "];"
End Function

Function searchArticles(code As String, name As String) As String
    searchArticles = "EXEC ('SELECT arvcexr, arccode, arvcexv, pkstrucobj.get_desc(123, arccinr, ''HR'') opis " _
            & "FROM artcoca, artuv WHERE arccinv = arvcinv"
       
    If Len(code) > 0 Then
        searchArticles = searchArticles & " AND (ARVCEXR like ''" & UCase(code) & "'' OR ARCCODE like ''" & UCase(code) & "'')"
    End If
    
    If Len(name) > 0 Then
        searchArticles = searchArticles & " AND pkstrucobj.get_desc(123, arccinr, ''HR'') like ''" & UCase(name) & "'' "
    End If
            
    searchArticles = searchArticles & "') at [" + db.getOracleServer + "];"
End Function

Function selectNtars() As String
    selectNtars = "EXEC ('SELECT DISTINCT (tventar), tvendesc FROM tra_avetar WHERE SUBSTR(tvendesc, 0, 2) NOT IN (''13'', ''11'') " _
                & "AND tventar > 1000 ORDER BY ( CASE " _
                & "WHEN SUBSTR(tventar, 6, 4) < 1000 THEN TO_NUMBER(SUBSTR(tventar, 7, 3) || SUBSTR(tventar, 0, 2))" _
                & "ELSE TO_NUMBER(''9'' || tventar)" _
                & "END ) DESC') at [" + db.getOracleServer + "];"
End Function

Function selectPrices(ntar As String, arvcexr As String, msnode As String, datum As String) As String

    selectPrices = "EXEC [" + db.getDatabase + "].[" + db.getProcedurePrefix + "].[" + db.getProcedure + "] @datum = '" & datum & "'"
    selectPrices = selectPrices + ", @site = NULL "
    
    If Len(ntar) = 0 Then
        selectPrices = selectPrices + ", @ntar = NULL "
    Else
        selectPrices = selectPrices + ", @ntar = N'" & ntar & "'"
    End If
    
    If Len(msnode) = 0 Then
        selectPrices = selectPrices + ", @msnode = N'-1'"
    Else
        selectPrices = selectPrices + ", @msnode = N'" & msnode & "'"
    End If
    
    If Len(arvcexr) = 0 Then
        selectPrices = selectPrices + ", @arvcexr = NULL "
    Else
        selectPrices = selectPrices + ", @arvcexr = N'" & arvcexr & "'"
    End If
    
End Function

Function selectFich() As String
    selectFich = "EXEC ('select asi_seq_util.nextval from dual') at [" + db.getOracleServer + "];"
End Function

Function killPrice(kodCjenika As String, datumPocetkaCijene As Date, datumKrajaCijene As Date, cijena As String, sifraArtikla As String, cexv As String, kodPoreza As String, fich As String, valuta As String)
    If datumKrajaCijene > Date Then
        If (datumPocetkaCijene > Date) Then
            killPrice = insertKillFuturePrice(kodCjenika, utils.getDateString(datumPocetkaCijene), utils.getDateString(datumKrajaCijene), sifraArtikla, Replace(cijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
            globals.addRowNumber
        Else
            killPrice = insertEndOfExistingPrice(kodCjenika, utils.getDateString(datumPocetkaCijene), utils.getDateString(Date), sifraArtikla, Replace(cijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
            globals.addRowNumber
        End If
    End If
    
End Function

Private Function insertEndOfExistingPrice(kodCjenika As Variant, datumPocetka As String, datumKraja As String, sifraArtikla As String, cijenaArtikla As String, cexv As String, kodPoreza As String, brojFichLinija As String, trenutniDatum As String, korisnik As String, fich As String, brojLinije As String, valuta As String) As String
    insertEndOfExistingPrice = "EXEC ('INSERT INTO intprixv (PVFNTAR, PVFIDDEB, PVFIDFIN, PVFCEXR, PVFCEXV, PVFPRIX, PVFCTVA, PVFRESID, PVFFLAG, PVFACTI, PVFLGFI, PVFTRT, PVFDTRT, PVFDCRE, PVFDMAJ, PVFUTIL, PVFFICH, PVFNLIG, PVFDEVI, PVFLANGUE, PVFMALIG) VALUES" _
                        & "(''" + kodCjenika + "'', " _
                        & "to_date(''" + datumPocetka + "'', ''yyyyMMdd''), " _
                        & "to_date(''" + datumKraja + "'', ''yyyyMMdd''), " _
                        & "''" + sifraArtikla + "'', " _
                        & "''" + cexv + "'', " _
                        & "''" + cijenaArtikla + "'', " _
                        & "''" + kodPoreza + "'', " _
                        & "''1'', " _
                        & "''3'', " _
                        & "''1'', " _
                        & "''" + brojFichLinija + "'', " _
                        & "''0''," _
                        & "to_date(''" + trenutniDatum + "'', ''yyyyMMdd''), " _
                        & "to_date(''" + trenutniDatum + "'', ''yyyyMMdd''), " _
                        & "to_date(''" + trenutniDatum + "'', ''yyyyMMdd''), " _
                        & "''" + Left(korisnik, 12) + "'', " _
                        & "''" + fich + "'', " _
                        & "" + brojLinije + ", " _
                        & "''" + valuta + "'', " _
                        & "''GB'', 0)') at [" + db.getOracleServer + "]; "
End Function

Private Function insertKillFuturePrice(kodCjenika As Variant, datumPocetka As String, datumKraja As String, sifraArtikla As String, cijenaArtikla As String, cexv As String, kodPoreza As String, brojFichLinija As String, trenutniDatum As String, korisnik As String, fich As String, brojLinije As String, valuta As String) As String
    'PVFACTI = 2 za brisanje
    insertKillFuturePrice = "EXEC ('INSERT INTO intprixv (PVFNTAR, PVFIDDEB, PVFIDFIN, PVFCEXR, PVFCEXV, PVFPRIX, PVFCTVA, PVFRESID, PVFFLAG, PVFACTI, PVFLGFI, PVFTRT, PVFDTRT, PVFDCRE, PVFDMAJ, PVFUTIL, PVFFICH, PVFNLIG, PVFDEVI, PVFLANGUE, PVFMALIG) VALUES" _
                        & "(''" + kodCjenika + "'', " _
                        & "to_date(''" + datumPocetka + "'', ''yyyyMMdd''), " _
                        & "to_date(''" + datumKraja + "'', ''yyyyMMdd''), " _
                        & "''" + sifraArtikla + "'', " _
                        & "''" + cexv + "'', " _
                        & "''" + cijenaArtikla + "'', " _
                        & "''" + kodPoreza + "'', " _
                        & "''1'', " _
                        & "''3'', " _
                        & "''2'', " _
                        & "''" + brojFichLinija + "'', " _
                        & "''0''," _
                        & "to_date(''" + trenutniDatum + "'', ''yyyyMMdd''), " _
                        & "to_date(''" + trenutniDatum + "'', ''yyyyMMdd''), " _
                        & "to_date(''" + trenutniDatum + "'', ''yyyyMMdd''), " _
                        & "''" + Left(korisnik, 12) + "'', " _
                        & "''" + fich + "'', " _
                        & "" + brojLinije + ", " _
                        & "''" + valuta + "'', " _
                        & "''GB'', 0)') at [" + db.getOracleServer + "]; "
End Function
