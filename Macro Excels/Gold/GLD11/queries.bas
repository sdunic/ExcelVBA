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


Function getHeader(orderNumber As String)
    getHeader = "EXEC ('"
    getHeader = getHeader & "SELECT ECCCEXCDE, to_char(ECCDCOM, ''DD.MM.RRRR''), ECCSITE, ECCNCLI,  pkclienctr.get_NUMCONTRAT(123, ECCNINT), ECCNFILF, "
    getHeader = getHeader & "adradre,  adrrais,  adrrue1,  adrcode || '' '' || adrvill,  pkparpostes.get_postlibc(123, 0, 807, adrpays, ''HR''), "
    getHeader = getHeader & "pkparpostes.get_postlibc(123, 0, 805, eccdevi, ''HR''),  pkparpostes.get_postlibc(123, 0, 950, NVL(ecccons, 0), ''HR''), to_char(eccdliv, ''DD.MM.RRRR''),  eccntou,  pkparpostes.get_postlibc(123, 0, 2015, eccetat, ''HR''), "
    getHeader = getHeader & "eccCOMM1 komentar,  eccnego global_discunt "
    getHeader = getHeader & "FROM cclentcde, clienctr, cliadres, clifilie "
    getHeader = getHeader & "WHERE ECCCEXCDE = ''" & orderNumber & "'' "
    getHeader = getHeader & "AND cclnint = eccnint AND adrncli = eccncli AND cfinfilc = eccnfilc AND cfincli = eccncli AND adradre = cficomm"
    getHeader = getHeader & "') at [" + db.getOracleServer + "];"
End Function

Function getDetails(orderNumber As String)
    getDetails = "EXEC ('"
    getDetails = getDetails & "SELECT dcccode, pkstrucobj.get_desc(123, aruclibl, ''HR''), arlcexvl, pkparpostes.get_postlibc(123, 0, 731, arutypul, ''HR''), PKTVAS.getTauxTVA(123, dccctva, 3, to_char(eccdcom,''DD/MM/RR'')), "
    getDetails = getDetails & "dccqtei / dccuauvc, dccuauvc, dccqtei, dccprfa, pkparpostes.get_postlibc(123, 0, 806, dccuapp, ''HR''), "
    getDetails = getDetails & "DECODE(pkparpostes.get_postvan1(123, 0, 806, dccuapp), 731, dccprfa * dccqtei / dccuauvc * dccuaut, dccprfa * dccqtei * dccuaut * dccpcu) "
    getDetails = getDetails & "FROM ccldetcde, artvl, artul, artrac, cclentcde "
    getDetails = getDetails & "WHERE dcccincde = ecccincde and ECCCEXCDE = ''" & orderNumber & "'' AND arlseqvl = dccseqvl AND dcccinl = arucinl AND dcccinr = artcinr "
    getDetails = getDetails & "') at [" + db.getOracleServer + "];"
End Function

Function getFooter(orderNumber As String)
    getFooter = "EXEC ('"
    getFooter = getFooter & "SELECT pccvale, pccbase, nvl(pccmont, 0) FROM cclpied, cclentcde "
    getFooter = getFooter & "WHERE pcccincde = ecccincde AND ECCCEXCDE = ''" & orderNumber & "'' "
    getFooter = getFooter & "AND pccrubr = 4 AND pcctdoc = 1 "
    getFooter = getFooter & "ORDER BY pccvale DESC "
    getFooter = getFooter & "') at [" + db.getOracleServer + "];"
End Function
