Attribute VB_Name = "queries"
Function selectMSGID() As String
    selectMSGID = "EXEC ('select asi_seq_msgid.nextval from dual') at [" + db.getOracleServer + "];"
End Function

Function insertASISTATUS(msgid As String) As String
    insertASISTATUS = "EXEC ('" _
    & "INSERT INTO asi_status (SASMSGID, SASAPPLI, SASSSTAT, SASSREAD, SASSPROC, SASSERR, SASGSTAT, SASGREAD, SASGPROC, SASGERR, SASDCRE, SASDMAJ, SASNMAJ, SASUTIL) " _
    & "VALUES (''" + msgid + "'', ''SIS_15'', 0, 0, 0, 0, 0, 0, 0, 0, sysdate, sysdate, 0, ''gb'')') at [" + db.getOracleServer + "];"
End Function

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


Function searchLocations(code As String, name As String) As String
    searchLocations = "EXEC ('" _
            & "SELECT robid, pkresobj.get_desc(123, robid, ''HR'') " _
            & "FROM resobj " _
            & "WHERE robid > 0 and robresid = ''1'' "
        
    If Len(code) > 0 Then
        searchLocations = searchLocations & " AND robid like ''" & UCase(code) & "'' "
    End If
    
    If Len(name) > 0 Then
        searchLocations = searchLocations & " AND pkresobj.get_desc(123, robid, ''HR'') like ''" & UCase(name) & "'' "
    End If
            
    searchLocations = searchLocations & "ORDER BY 1') at [" + db.getOracleServer + "];"
End Function

Function searchCustomers(code As String, name As String) As String
    searchCustomers = "EXEC ('SELECT clincli, TRIM(clilibl) " _
            & "FROM clidgene WHERE clityma = 1 "
            
    If Len(code) > 0 Then
        searchCustomers = searchCustomers & " AND clincli like ''" & UCase(code) & "'' "
    End If
    
    If Len(name) > 0 Then
        searchCustomers = searchCustomers & " AND clilibl like ''" & UCase(name) & "'' "
    End If
          
    searchCustomers = searchCustomers & " ORDER BY 2')"
    searchCustomers = searchCustomers & "at [" + db.getOracleServer + "];"
End Function

Function searchContracts(code As String, name As String, customer As String, searchCustomer As Boolean) As String
    
    searchContracts = "EXEC ('" _
            & "SELECT DISTINCT cclnum, clincli, clilibl FROM clidgene, cliadres, " _
            & "(SELECT CCLNUM, LCCNCLI FROM lienconcli, clienctr WHERE CCLNINT = LCCNINT AND TRUNC(current_date) BETWEEN lccddeb AND lccdfin) " _
            & "WHERE clincli = ADRNCLI AND LCCNCLI(+) = ADRNCLI AND CCLNUM IS NOT NULL " _

    If Len(code) > 0 Then
        searchContracts = searchContracts & "AND (cclnum like ''" & UCase(code) & "'') "
    End If
    
    If Len(name) > 0 Then
        searchContracts = searchContracts & "AND clilibl like ''" & UCase(name) & "'' "
    End If
    
    If Len(customer) > 0 And searchCustomer = True Then
        searchContracts = searchContracts & "AND clincli = ''" & UCase(Split(customer, " | ")(0)) & "'' "
    End If
    
    searchContracts = searchContracts & "ORDER BY cclnum ASC') at [" + db.getOracleServer + "];"

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
            & "FROM wpline, wplig, artcoca, artuv WHERE wlgcinl = arvcinv and wlgcinl = arccinv and wlgcinwpl = wplcinwpl and wplnum = ''CORE-NON'' "
       
    If Len(code) > 0 Then
        searchArticles = searchArticles & " AND (ARVCEXR like ''%" & UCase(code) & "%'' OR ARCCODE like ''%" & UCase(code) & "%'')"
    End If
    
    If Len(name) > 0 Then
        searchArticles = searchArticles & " AND pkstrucobj.get_desc(123, arccinr, ''HR'') like ''" & UCase(name) & "'' "
    End If
            
    searchArticles = searchArticles & "') at [" + db.getOracleServer + "];"
End Function


Function getCexrs() As String
    getCexrs = "EXEC ('SELECT arvcexr FROM wpline, wplig, artcoca, artuv WHERE wlgcinl = arvcinv and wlgcinl = arccinv and wlgcinwpl = wplcinwpl and wplnum = ''CORE-NON'' "
    getCexrs = getCexrs & "') at [" + db.getOracleServer + "];"
End Function


Function searchAnalyticalArticles(code As String, name As String) As String
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


Function insertSIS15(msgid As String, nlig As String, site As String, tipFakture As String, kupac As String, ugovor As String, _
datumFakture As String, cexr As String, cexrAnaliza As String, msAnaliza As String, tm_pc As String, reasonCode As String, korisnik As String, _
napomena As String, kolicinaStavke As Double, iznosStavke As Double, lv As String, aSite As String, aNW As String, TCVFYP As String) As String

    insertSIS15 = "EXEC ('"
    insertSIS15 = insertSIS15 & "INSERT INTO sis15_core_inv (TCVMSGID, TCVLNLIG, TCVSITE, TCVSUM, TCVNCLI, TCVCNUM, TCVDATEF, TCVCEXR, TCVLV, TCVLU, "
    insertSIS15 = insertSIS15 & "TCVPU, TCVMONT, TCVTM, TCVMOTF, TCVMS, TCVCEXRA, TCVNOTE, TCVUSER, TCVSTRT, TCVSDTRT, TCVGTRT, TCVGDTRT, TCVDCRE, TCVDMAJ, TCVTIL, TCVASITE, TCVANW, TCVFYP)"
    insertSIS15 = insertSIS15 & " VALUES ("
    
    insertSIS15 = insertSIS15 & msgid & ", " 'TCVMSGID NUMBER(9,0)
    insertSIS15 = insertSIS15 & nlig & ", " 'TCVLNLIG  NUMBER(9,0)
    insertSIS15 = insertSIS15 & site & ", " 'TCVSITE   NUMBER(5,0)
    insertSIS15 = insertSIS15 & tipFakture & ", " 'TCVSUM  NUMBER(1,0)
    insertSIS15 = insertSIS15 & "''" & kupac & "'', " 'TCVNCLI VARCHAR2(9 CHAR)
    insertSIS15 = insertSIS15 & "''" & ugovor & "'', " 'TCVCNUM    VARCHAR2(8 CHAR)
    insertSIS15 = insertSIS15 & utils.getDateString(CDate(datumFakture)) & ", " 'TCVDATEF  DATE
    insertSIS15 = insertSIS15 & "''" & cexr & "'', " 'TCVCEXR  VARCHAR2(13 CHAR)
    insertSIS15 = insertSIS15 & lv & ", " 'TCVLV   NUMBER(9,0)
    insertSIS15 = insertSIS15 & "1, " 'TCVLU   NUMBER(9,0)
    insertSIS15 = insertSIS15 & utils.getPriceValue(kolicinaStavke) & ", " 'TCVPU   NUMBER(9,3)
    insertSIS15 = insertSIS15 & utils.getPriceValue(iznosStavke) & ", " 'TCVMONT    NUMBER(15,5)
    insertSIS15 = insertSIS15 & "''" & tm_pc & "'', " 'TCVTM   NUMBER(9,0)
    insertSIS15 = insertSIS15 & reasonCode & ", " 'TCVMOTF   NUMBER(3,0)
    
    If Len(msAnaliza) > 0 Then
        insertSIS15 = insertSIS15 & "''" & Split(msAnaliza, " | ")(0) & "'', " 'TCVMS   VARCHAR2(13 CHAR)
    Else
        insertSIS15 = insertSIS15 & "null, "
    End If
    
    If Len(cexrAnaliza) > 0 Then
        insertSIS15 = insertSIS15 & "''" & Split(cexrAnaliza, " | ")(0) & "'', " 'TCVCEXRA  VARCHAR2(13 CHAR)
    Else
        insertSIS15 = insertSIS15 & "null, "
    End If
    
    insertSIS15 = insertSIS15 & "''" & napomena & "'', " 'TCVNOTE  VARCHAR2(160 CHAR)
    insertSIS15 = insertSIS15 & "''" & korisnik & "'', " 'TCVUSER  VARCHAR2(320 CHAR)
    insertSIS15 = insertSIS15 & "0, " 'TCVSTRT NUMBER(3,0)
    insertSIS15 = insertSIS15 & "sysdate, " 'TCVSDTRT  DATE
    insertSIS15 = insertSIS15 & "0, " 'TCVGTRT NUMBER(3,0)
    insertSIS15 = insertSIS15 & "sysdate, " 'TCVGDTRT  DATE
    insertSIS15 = insertSIS15 & "sysdate, " 'TCVDCRE   DATE
    insertSIS15 = insertSIS15 & "sysdate, " 'TCVDMAJ   DATE
    insertSIS15 = insertSIS15 & "''xlsx_fact'', " 'TCVTIL    VARCHAR2(12 CHAR)
    
    If Len(aSite) = 0 Then
        aSite = "NULL"
    End If
    insertSIS15 = insertSIS15 & aSite & ", " 'TCVASITE   NUMBER(5,0)
    
    If Len(aNW) = 0 Then
        aNW = "NULL"
    End If
    insertSIS15 = insertSIS15 & aNW & ", "  'TCVANW   NUMBER(5,0)
    
    If Len(TCVFYP) = 0 Then
        TCVFYP = "NULL"
    End If
    insertSIS15 = insertSIS15 & TCVFYP 'TCVFYP  - 2 oznaka storno
    
    insertSIS15 = insertSIS15 & " )"
    insertSIS15 = insertSIS15 & "') at [" + db.getOracleServer + "];"

End Function


Function getStornoData(brojFakture As String, datumFakture As Date) As String

    getStornoData = "EXEC ('"
    getStornoData = getStornoData + "select "
    getStornoData = getStornoData + "   (SELECT DECODE(tcvmotf, 901, ''901 | CORE FAKTURA'', 902, ''902 | NON CORE FAKTURA'', null) FROM SIS15_CORE_INV WHERE TCVMSGID = TVCMSGID and TCVLNLIG = 1) FAC_MOTF, "
    getStornoData = getStornoData + "   TVCSITE || '' | '' || (select max(soclmag) from sitdgene where socsite = TVCSITE) FAC_SITE, "
    getStornoData = getStornoData + "   (SELECT tcvsum FROM SIS15_CORE_INV WHERE TCVMSGID = TVCMSGID and TCVLNLIG = 1) FAC_SUM,  "
    getStornoData = getStornoData + "   (SELECT tcvncli || '' | '' || (select clilibl from clidgene where clincli = tcvncli) FROM SIS15_CORE_INV WHERE TCVMSGID = TVCMSGID and TCVLNLIG = 1) FAC_NCLI, "
    getStornoData = getStornoData + "   (SELECT tcvcnum || '' | '' || (select clilibl from clidgene where clincli = tcvncli) FROM SIS15_CORE_INV WHERE TCVMSGID = TVCMSGID and TCVLNLIG = 1) FAC_CNUM, "
    getStornoData = getStornoData + "   to_date(current_date) FAC_DATF, "
    getStornoData = getStornoData + "   (select arvcexr || '' | '' || pkartcoca.get_closestean(123, arvcinv) || '' | '' || pkstrucobj.get_desc(123, arvcinr, ''HR'') || '' | '' || arvcexv from artuv where arvcinv = TVCCINL) FAC_ARTICLE,  "
    getStornoData = getStornoData + "   (select arvcexv || '' | SKU'' from artuv where arvcinv = TVCCINL) FAC_LV, "
    getStornoData = getStornoData + "   (TVCQTE) FAC_QTE, "
    getStornoData = getStornoData + "   (TVCMONT) FAC_MONT, "
    getStornoData = getStornoData + "   TVCANASITE FAC_TM, "
    getStornoData = getStornoData + "   TVCANACEXR FAC__A_CEXR, "
    getStornoData = getStornoData + "   TVCANAMS FAC__A_MS, "
    getStornoData = getStornoData + "   TVCASITE FAC__A_SITE, "
    getStornoData = getStornoData + "   TVCANW FAC__A_NW, "
    getStornoData = getStornoData + "   TVCNLIG FAC_NLIG "
    getStornoData = getStornoData + "from tom_core_inv "
    getStornoData = getStornoData + "where TVCCINPRO in ( "
    getStornoData = getStornoData + "   select lffrcin from facentfac, faclien "
    getStornoData = getStornoData + "   where LFFCINFAC = effcinfac And lffrtyp = 4 And effsite = 99999 "
    getStornoData = getStornoData + "   and effcexcus = ''" + brojFakture + "'' "
    getStornoData = getStornoData + "   and effdatf = " + utils.getDateString(datumFakture) + " "
    getStornoData = getStornoData + ") order by TVCNLIG "
    getStornoData = getStornoData + "') at [" + db.getOracleServer + "];"

End Function


