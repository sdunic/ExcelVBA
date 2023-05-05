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

Function selectMSNodes() As String
    selectMSNodes = "EXEC ('" _
            & "SELECT DISTINCT pkstrucobj.get_cext(123, objcint) MS_CODE, pkstrucobj.get_desc(123, objcint, ''HR'') MS_DESC " _
            & "FROM strucrel " _
            & "WHERE NOT EXISTS (SELECT 1 FROM artrac WHERE artcinr = objcint) " _
            & "ORDER BY 1') at [" + db.getOracleServer + "];"
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


Function selectArticles() As String
    ' arvcexr - interna šifra artikla
    ' arccode - barkod artikla
    selectArticles = "EXEC ('SELECT arvcexr, arccode, arvcexv, pkstrucobj.get_desc(123, arccinr, ''HR'') opis " _
            & "FROM artcoca, artuv WHERE arccinv = arvcinv') at [" + db.getOracleServer + "];"
End Function

Function searchArticles(code As String, name As String) As String
    searchArticles = "EXEC ('SELECT arvcexr, arccode, arvcexv, pkstrucobj.get_desc(123, arccinr, ''HR'') opis " _
            & "FROM artcoca, artuv WHERE arccinv = arvcinv"
       
    If Len(code) > 0 Then
        searchArticles = searchArticles & " AND (ARVCEXR like ''%" & UCase(code) & "%'' OR ARCCODE like ''" & UCase(code) & "'')"
    End If
    
    If Len(name) > 0 Then
        searchArticles = searchArticles & " AND pkstrucobj.get_desc(123, arccinr, ''HR'') like ''" & UCase(name) & "'' "
    End If
            
    searchArticles = searchArticles & "') at [" + db.getOracleServer + "];"
End Function

Function selectSuppliers() As String
    ' foucnuf - eksterni kod dobavljaèa
    ' foulibl - naziv dobavljaèa
    ' foucfin - interna šifra dobavljaèa za upite
    selectSuppliers = "EXEC ('SELECT foucnuf, foulibl, foucfin " _
            & "FROM foudgene WHERE foutype = 1 ORDER BY 2') at [" + db.getOracleServer + "];"
End Function

Function searchSuppliers(code As String, name As String) As String
    searchSuppliers = "EXEC ('SELECT foucnuf, foulibl, foucfin " _
            & "FROM foudgene WHERE foutype = 1 "
            
    If Len(code) > 0 Then
        searchSuppliers = searchSuppliers & " AND foucnuf like ''" & UCase(code) & "'' "
    End If
    
    If Len(name) > 0 Then
        searchSuppliers = searchSuppliers & " AND foulibl like ''" & UCase(name) & "'' "
    End If
          
    searchSuppliers = searchSuppliers & " ORDER BY 2')"
    searchSuppliers = searchSuppliers & "at [" + db.getOracleServer + "];"
End Function


Function selectPrices(objcint As String, arvcexr As String, cfin As String) As String
    'objcint = gold šifra èvora, arvcexr = gold šifra artikla ili barkod
    'funkcija prima ili jedan ili drugi parametar
    'ako je barkod onda trebamo dohvatiti šifru artikla
    
    If Len(objcint) = 0 Then
        objcint = -1
    End If
    
    If Len(arvcexr) = 0 Then
        arvcexr = -1
    End If
    
    If Len(cfin) = 0 Then
        cfin = -1
    End If
    
    selectPrices = "EXEC [" + db.getDatabase + "].[" + db.getProcedurePrefix + "].[" + db.getProcedure + "] @objcint = N'" & objcint & "', @arvcexr = N'" & arvcexr & "', @cfin = N'" & cfin & "'"
End Function

Function selectFich() As String
    selectFich = "EXEC ('select asi_seq_util.nextval from dual') at [" + db.getOracleServer + "];"
End Function

Function getInsertPrix(kodCjenika As String, datumCijene As Date, staraCijena As String, novaCijena As String, sifraArtikla As String, cexv As String, kodPoreza As String, fich As String, valuta As String) As String
    If datumCijene > CDate("31-12-1899") And Len(novaCijena) > 0 Then
        getInsertPrix = getInsertPrix + queries.insertEndOfPrice(kodCjenika, utils.getDateString(datumCijene), utils.getDateString(Date), sifraArtikla, Replace(staraCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
        globals.addRowNumber
        If novaCijena > 0 Then
            getInsertPrix = getInsertPrix + queries.insertNewPrice(kodCjenika, utils.getDateString(Date + 1), "31-12-2049", sifraArtikla, Replace(novaCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
            globals.addRowNumber
        End If
    ElseIf Len(novaCijena) > 0 Then
        If CDbl(novaCijena) > 0 Then
            getInsertPrix = getInsertPrix + queries.insertNewPrice(kodCjenika, utils.getDateString(Date + 1), "31-12-2049", sifraArtikla, Replace(novaCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
            globals.addRowNumber
        End If
    End If
End Function

Private Function insertEndOfPrice(kodCjenika As Variant, datumPocetka As String, datumKraja As String, sifraArtikla As String, cijenaArtikla As String, cexv As String, kodPoreza As String, brojFichLinija As String, trenutniDatum As String, korisnik As String, fich As String, brojLinije As String, valuta As String) As String
    insertEndOfPrice = "EXEC ('INSERT INTO intprixv (PVFNTAR, PVFIDDEB, PVFIDFIN, PVFCEXR, PVFCEXV, PVFPRIX, PVFCTVA, PVFRESID, PVFFLAG, PVFACTI, PVFLGFI, PVFTRT, PVFDTRT, PVFDCRE, PVFDMAJ, PVFUTIL, PVFFICH, PVFNLIG, PVFDEVI, PVFLANGUE, PVFMALIG) VALUES" _
                        & "(''" + kodCjenika + "'', " _
                        & "to_date(''" + datumPocetka + "'', ''dd-mm-yyyy''), " _
                        & "to_date(''" + datumKraja + "'', ''dd-mm-yyyy''), " _
                        & "''" + sifraArtikla + "'', " _
                        & "''" + cexv + "'', " _
                        & "''" + cijenaArtikla + "'', " _
                        & "''" + kodPoreza + "'', " _
                        & "''1'', " _
                        & "''3'', " _
                        & "''1'', " _
                        & "''" + brojFichLinija + "'', " _
                        & "''0''," _
                        & "to_date(''" + trenutniDatum + "'', ''dd-mm-yyyy''), " _
                        & "to_date(''" + trenutniDatum + "'', ''dd-mm-yyyy''), " _
                        & "to_date(''" + trenutniDatum + "'', ''dd-mm-yyyy''), " _
                        & "''" + korisnik + "'', " _
                        & "''" + fich + "'', " _
                        & "" + brojLinije + ", " _
                        & "''" + valuta + "'', " _
                        & "''GB'', 0)') at [" + db.getOracleServer + "]; "
End Function


Private Function insertNewPrice(kodCjenika As Variant, datumPocetka As String, datumKraja As String, sifraArtikla As String, cijenaArtikla As String, cexv As String, kodPoreza As String, brojFichLinija As String, trenutniDatum As String, korisnik As String, fich As String, brojLinije As String, valuta As String) As String
    insertNewPrice = "EXEC ('INSERT INTO intprixv (PVFNTAR, PVFIDDEB, PVFIDFIN, PVFCEXR, PVFCEXV, PVFPRIX, PVFCTVA, PVFRESID, PVFFLAG, PVFACTI, PVFLGFI, PVFTRT, PVFDTRT, PVFDCRE, PVFDMAJ, PVFUTIL, PVFFICH, PVFNLIG, PVFDEVI, PVFLANGUE, PVFMALIG) VALUES" _
                        & "(''" + kodCjenika + "'', " _
                        & "to_date(''" + datumPocetka + "'', ''dd-mm-yyyy''), " _
                        & "to_date(''31-12-2049'', ''dd-mm-yyyy''), " _
                        & "''" + sifraArtikla + "'', " _
                        & "''" + cexv + "'', " _
                        & "''" + cijenaArtikla + "'', " _
                        & "''" + kodPoreza + "'', " _
                        & "''1'', " _
                        & "''3'', " _
                        & "''1'', " _
                        & "''" + brojFichLinija + "'', " _
                        & "''0''," _
                        & "to_date(''" + trenutniDatum + "'', ''dd-mm-yyyy''), " _
                        & "to_date(''" + trenutniDatum + "'', ''dd-mm-yyyy''), " _
                        & "to_date(''" + trenutniDatum + "'', ''dd-mm-yyyy''), " _
                        & "''" + korisnik + "'', " _
                        & "''" + fich + "'', " _
                        & "" + brojLinije + ", " _
                        & "''" + valuta + "'', " _
                        & "''GB'', 0)') at [" + db.getOracleServer + "]; "
End Function
