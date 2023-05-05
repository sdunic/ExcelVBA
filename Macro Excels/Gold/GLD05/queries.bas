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
        searchArticles = searchArticles & " AND (ARVCEXR like ''%" & UCase(code) & "%'' OR ARCCODE like ''%" & UCase(code) & "%'')"
    End If
    
    If Len(name) > 0 Then
        searchArticles = searchArticles & " AND pkstrucobj.get_desc(123, arccinr, ''HR'') like ''" & UCase(name) & "'' "
    End If
            
    searchArticles = searchArticles & "') at [" + db.getOracleServer + "];"
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

Function searchStores(code As String) As String
    searchStores = "EXEC ('SELECT aventar, tvendesc " _
            & " FROM avetar, tra_avetar, avescope, sitdgene " _
            & " WHERE aventar = tventar AND langue = ''HR'' AND avorescint = socsite AND avontar = aventar " _
            & " AND trunc(current_date) BETWEEN avoddeb AND avodfin HAVING count(1) = 1 " _
            & " AND aventar like ''" & code & "%'' " _
            & " GROUP BY aventar, tvendesc ORDER BY SUBSTR(AVENTAR, 5, 5), SUBSTR(AVENTAR, 0, 1)') at [" + db.getOracleServer + "];"
End Function

Function selectLocalPrices(ntar As Long, objcint As String, cfin As String, arvcexr As String, sites As String, datum As Date) As String
    'ntar = gold šifra lokalnog cjenika
    'objcint = gold šifra èvora
    'cfin = gold šifra dobavljaèa
    
    If Len(objcint) = 0 Then
        objcint = -1
    End If
    
    If ntar = 0 Then
        ntar = -1
    End If
    
    If Len(cfin) = 0 Then
        cfin = -1
    End If
    
    If Len(arvcexr) = 0 Then
        arvcexr = -1
    End If
    
    selectLocalPrices = "EXEC [" + db.getDatabase + "].[" + db.getProcedurePrefix + "].[" + db.getProcedure + "] @objcint = N'" & objcint & "', @arvcexr = N'" & arvcexr & "', @cfin = N'" & cfin & "', @ntarType = " & ntar & ", @sites = '" & sites & "', @datum = '" & utils.getDateStringProcedure(datum) & "'"
    
End Function

Function selectFich() As String
    selectFich = "EXEC ('select asi_seq_util.nextval from dual') at [" + db.getOracleServer + "];"
End Function

Function getInsertPrix(kodCjenika As String, datumStareCijene As Date, datumKrajaStareCijene As Date, staraCijena As String, _
    datumNoveCijene As Date, datumKrajaNoveCijene As Date, novaCijena As String, _
    sifraArtikla As String, cexv As String, kodPoreza As String, fich As String, valuta As String) As String
    
    If datumKrajaNoveCijene = CDate("0:00:00") Then
        datumKrajaNoveCijene = CDate("31-12-2049")
    End If
    
    If staraCijena = 0 Or datumNoveCijene > datumKrajaStareCijene Then
        getInsertPrix = getInsertPrix + queries.insertPrice(kodCjenika, utils.getDateString(datumNoveCijene), utils.getDateString(datumKrajaNoveCijene), sifraArtikla, _
        Replace(novaCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
        globals.addRowNumber
    Else
        'ako je stara cijena u relativnoj buduænosti naspram nove cijene, staru cijenu skroz gasimo kao da nije postojala
        'ako je stara cijena u relativnoj prošlosti naspram nove cijene, staru cijenu gasimo dan ranije od poèetka nove cijene
        If (datumStareCijene > Date) Then
            getInsertPrix = getInsertPrix + insertKillFuturePrice(kodCjenika, utils.getDateString(datumStareCijene), utils.getDateString(datumKrajaStareCijene), sifraArtikla, _
            Replace(staraCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
            globals.addRowNumber
        Else
            getInsertPrix = getInsertPrix + insertEndOfExistingPrice(kodCjenika, utils.getDateString(datumStareCijene), utils.getDateString(Date), sifraArtikla, _
            Replace(staraCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
            globals.addRowNumber
        End If
    
        'nakon "zatvaranja" cijene, palimo nove cijene u kombinaciji sa starim cijenama
        If (datumNoveCijene <= datumStareCijene) Then
            If (datumKrajaNoveCijene >= datumKrajaStareCijene) Then
                getInsertPrix = getInsertPrix + insertPrice(kodCjenika, utils.getDateString(datumNoveCijene), utils.getDateString(datumKrajaNoveCijene), sifraArtikla, _
                Replace(novaCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
                globals.addRowNumber
            
            ElseIf (datumKrajaNoveCijene < datumPocetkaStareCijene) Then
                getInsertPrix = getInsertPrix + insertPrice(kodCjenika, utils.getDateString(datumNoveCijene), utils.getDateString(datumKrajaNoveCijene), sifraArtikla, _
                Replace(novaCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
                globals.addRowNumber
                getInsertPrix = getInsertPrix + insertPrice(kodCjenika, utils.getDateString(datumStareCijene), utils.getDateString(datumKrajaStareCijene), sifraArtikla, _
                Replace(staraCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
                globals.addRowNumber
            
            Else
                getInsertPrix = getInsertPrix + insertPrice(kodCjenika, utils.getDateString(datumNoveCijene), utils.getDateString(datumKrajaNoveCijene), sifraArtikla, _
                Replace(novaCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
                globals.addRowNumber
                getInsertPrix = getInsertPrix + insertPrice(kodCjenika, utils.getDateString(datumKrajaNoveCijene + 1), utils.getDateString(datumKrajaStareCijene), sifraArtikla, _
                Replace(staraCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
                globals.addRowNumber
                
            End If
        ElseIf (datumNoveCijene > datumStareCijene And datumNoveCijene <= datumKrajaStareCijene) Then
            If (datumKrajaNoveCijene >= datumKrajaStareCijene) Then
                getInsertPrix = getInsertPrix + insertPrice(kodCjenika, utils.getDateString(datumStareCijene), utils.getDateString(datumNoveCijene - 1), sifraArtikla, _
                Replace(staraCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
                globals.addRowNumber
                getInsertPrix = getInsertPrix + insertPrice(kodCjenika, utils.getDateString(datumNoveCijene), utils.getDateString(datumKrajaNoveCijene), sifraArtikla, _
                Replace(novaCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
                globals.addRowNumber
            Else
                getInsertPrix = getInsertPrix + insertPrice(kodCjenika, utils.getDateString(datumStareCijene), utils.getDateString(datumNoveCijene - 1), sifraArtikla, _
                Replace(staraCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
                globals.addRowNumber
                getInsertPrix = getInsertPrix + insertPrice(kodCjenika, utils.getDateString(datumNoveCijene), utils.getDateString(datumKrajaNoveCijene), sifraArtikla, _
                Replace(novaCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
                globals.addRowNumber
                getInsertPrix = getInsertPrix + insertPrice(kodCjenika, utils.getDateString(datumKrajaNoveCijene + 1), utils.getDateString(datumKrajaStareCijene), sifraArtikla, _
                Replace(staraCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
                globals.addRowNumber
            
            End If
        
        Else
            getInsertPrix = getInsertPrix + insertPrice(kodCjenika, utils.getDateString(datumNoveCijene), utils.getDateString(datumKrajaNoveCijene), sifraArtikla, _
            Replace(novaCijena, ",", "."), cexv, kodPoreza, globals.getRowCount, utils.getDateString(Date), utils.getUserName, fich, globals.getRowNumber, valuta)
            globals.addRowNumber
        End If
  
    End If
    
End Function

Private Function insertPrice(kodCjenika As Variant, datumPocetka As String, datumKraja As String, sifraArtikla As String, cijenaArtikla As String, cexv As String, kodPoreza As String, brojFichLinija As String, trenutniDatum As String, korisnik As String, fich As String, brojLinije As String, valuta As String) As String
    insertPrice = "EXEC ('INSERT INTO intprixv (PVFNTAR, PVFIDDEB, PVFIDFIN, PVFCEXR, PVFCEXV, PVFPRIX, PVFCTVA, PVFRESID, PVFFLAG, PVFACTI, PVFLGFI, PVFTRT, PVFDTRT, PVFDCRE, PVFDMAJ, PVFUTIL, PVFFICH, PVFNLIG, PVFDEVI, PVFLANGUE, PVFMALIG) VALUES" _
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
                        & "''" + Left(korisnik, 12) + "'', " _
                        & "''" + fich + "'', " _
                        & "" + brojLinije + ", " _
                        & "''" + valuta + "'', " _
                        & "''GB'', 0)') at [" + db.getOracleServer + "]; "
End Function

Private Function insertEndOfExistingPrice(kodCjenika As Variant, datumPocetka As String, datumKraja As String, sifraArtikla As String, cijenaArtikla As String, cexv As String, kodPoreza As String, brojFichLinija As String, trenutniDatum As String, korisnik As String, fich As String, brojLinije As String, valuta As String) As String
    insertEndOfExistingPrice = "EXEC ('INSERT INTO intprixv (PVFNTAR, PVFIDDEB, PVFIDFIN, PVFCEXR, PVFCEXV, PVFPRIX, PVFCTVA, PVFRESID, PVFFLAG, PVFACTI, PVFLGFI, PVFTRT, PVFDTRT, PVFDCRE, PVFDMAJ, PVFUTIL, PVFFICH, PVFNLIG, PVFDEVI, PVFLANGUE, PVFMALIG) VALUES" _
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
                        & "to_date(''" + datumPocetka + "'', ''dd-mm-yyyy''), " _
                        & "to_date(''" + datumKraja + "'', ''dd-mm-yyyy''), " _
                        & "''" + sifraArtikla + "'', " _
                        & "''" + cexv + "'', " _
                        & "''" + cijenaArtikla + "'', " _
                        & "''" + kodPoreza + "'', " _
                        & "''1'', " _
                        & "''3'', " _
                        & "''2'', " _
                        & "''" + brojFichLinija + "'', " _
                        & "''0''," _
                        & "to_date(''" + trenutniDatum + "'', ''dd-mm-yyyy''), " _
                        & "to_date(''" + trenutniDatum + "'', ''dd-mm-yyyy''), " _
                        & "to_date(''" + trenutniDatum + "'', ''dd-mm-yyyy''), " _
                        & "''" + Left(korisnik, 12) + "'', " _
                        & "''" + fich + "'', " _
                        & "" + brojLinije + ", " _
                        & "''" + valuta + "'', " _
                        & "''GB'', 0)') at [" + db.getOracleServer + "]; "
End Function


