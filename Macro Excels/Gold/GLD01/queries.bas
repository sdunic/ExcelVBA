Attribute VB_Name = "queries"
Function selectFich() As String
    selectFich = "EXEC ('select asi_seq_util.nextval from dual') at [" + db.getOracleServer + "];"
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

Function selectLocations() As String
    selectLocations = "EXEC ('" _
            & "SELECT robid, pkresobj.get_desc(123, robid, ''HR'') " _
            & "FROM resobj " _
            & "WHERE robresid = ''1''" _
            & "ORDER BY 1') at [" + db.getOracleServer + "];"
End Function
Function searchLocations(code As String, name As String) As String
    searchLocations = "EXEC ('" _
            & "SELECT robid, pkresobj.get_desc(123, robid, ''HR'') " _
            & "FROM resobj " _
            & "WHERE robresid = ''1'' "
        
    If Len(code) > 0 Then
        searchLocations = searchLocations & " AND robid like ''" & UCase(code) & "'' "
    End If
    
    If Len(name) > 0 Then
        searchLocations = searchLocations & " AND pkresobj.get_desc(123, robid, ''HR'') like ''" & UCase(name) & "'' "
    End If
            
    searchLocations = searchLocations & "ORDER BY 1') at [" + db.getOracleServer + "];"
End Function

Function selectContracts(val As Integer) As String
    'val - interni kod dobavljaèa
    selectContracts = "EXEC ('" _
            & "SELECT distinct fccnum, fcclib " _
            & "FROM fouccom, liencom " _
            & "WHERE licccin = fccccin " _
            & "AND (liccfin = " & val & " or ''-1'' = " & val & ") " _
            & "ORDER BY fccnum ASC') at [" + db.getOracleServer + "];"

End Function
Function searchContracts(code As String, name As String, supplier As String, searchSupplier As Boolean) As String
    'val - interni kod dobavljaèa
    searchContracts = "EXEC ('" _
            & "SELECT distinct fccnum, fcclib " _
            & "FROM fouccom, liencom " _
            & "WHERE licccin = fccccin "
        
    If Len(code) > 0 Then
        searchContracts = searchContracts & "AND (fccnum like ''" & UCase(code) & "'') "
    End If
    
    If Len(name) > 0 Then
        searchContracts = searchContracts & "AND fcclib like ''" & UCase(name) & "'' "
    End If
    
    If Len(supplier) > 0 And searchSupplier = True Then
        searchContracts = searchContracts & "AND liccfin = ''" & UCase(Split(supplier, " - ")(1)) & "'' "
    End If
    
    searchContracts = searchContracts & "ORDER BY fccnum ASC') at [" + db.getOracleServer + "];"

End Function

Function selectArticleLists() As String
    selectArticleLists = "EXEC ('" _
            & "SELECT elinlis, NVL(elilibl, '' '') " _
            & "FROM artentlist " _
            & "') at [" + db.getOracleServer + "];"
End Function
Function searchArticleLists(code As String, name As String) As String
    searchArticleLists = "EXEC ('" _
            & "SELECT elinlis, NVL(elilibl, '' '') " _
            & "FROM artentlist "
    
    If Len(code) > 0 Then
        searchArticleLists = searchArticleLists & "WHERE (elinlis like ''" & UCase(code) & "'') "
    End If
    
    If Len(name) > 0 Then
        searchArticleLists = searchArticleLists & "WHERE NVL(elilibl, '' '') like ''" & UCase(name) & "'' "
    End If
    
    searchArticleLists = searchArticleLists & "') at [" + db.getOracleServer + "];"
End Function

Function selectArticleGroup() As String
    selectArticleGroup = "EXEC ('" _
            & "SELECT distinct tgacexgat, tgalibl " _
            & "FROM targrar " _
            & "') at [" + db.getOracleServer + "];"
End Function
Function searchArticleGroups(code As String, name As String) As String
    searchArticleGroups = "EXEC ('" _
            & "SELECT distinct tgacexgat, tgalibl " _
            & "FROM targrar "
    
    If Len(code) > 0 Then
        searchArticleGroups = searchArticleGroups & "WHERE (tgacexgat like ''" & UCase(code) & "'') "
    End If
    
    If Len(name) > 0 Then
        searchArticleGroups = searchArticleGroups & "WHERE tgalibl like ''" & UCase(name) & "'' "
    End If
    
    searchArticleGroups = searchArticleGroups & "') at [" + db.getOracleServer + "];"
End Function


Function selectArticles() As String
    selectArticles = "EXEC ('" _
            & "SELECT artcexr, pkstrucobj.get_desc(123,artcinr,''HR'') " _
            & "FROM artrac" _
            & "') at [" + db.getOracleServer + "];"
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

Function selectClasses() As String
    selectClasses = "EXEC ('" _
            & "SELECT tatccode, tatclibc " _
            & "FROM tra_attricla WHERE langue =''HR'' " _
            & "ORDER BY tatccode " _
            & "') at [" + db.getOracleServer + "];"
End Function

Function searchClassAttributes(class As String) As String
    searchClassAttributes = "EXEC ('" _
            & "SELECT tattcode, tattlibl " _
            & "FROM tra_attrival WHERE langue=''HR'' and TATTCCLA = ''" & class & "'' " _
            & "ORDER BY tattcode') at [" + db.getOracleServer + "];"
End Function


Function loadPurchaseConditionsDataToInterfaceTable(num_log As Integer, cnuf As String, cnum As String, indate As String, art_grp As String, cexr As String, site As Integer, ms As String, ccla As String, catt As String, text As String, num As Integer, artlist As String, domain_user As String) As String
    loadPurchaseConditionsDataToInterfaceTable = "BEGIN DECLARE @message varchar(500), @nb_lines int, @msgid int, @reutrn int " _
            & "EXEC ('" _
            & "DECLARE " _
            & "o_message varchar2(500); o_nb_lines number; o_msgid number; o_reutrn number; " _
            & "BEGIN " _
            & "PKTMY_PURCH_COND.GENERATE_DATA_IN_sis14_nab_uvjeti( " _
            & num_log & ", " _
            & "''" & cnuf & "'', " _
            & "''" & cnum & "'', " _
            & "''" & indate & "'', " _
            & "''" & art_grp & "'', " _
            & "''" & cexr & "'', " _
            & site & ", " _
            & "''" & ms & "'', " _
            & "''" & ccla & "'', " _
            & "''" & catt & "'', " _
            & "''" & text & "'', " _
            & num & ", " _
            & "''" & artlist & "'', " _
            & "''" & domain_user & "'', " _
            & "?, ?, ?, ? ); END;', @message OUTPUT, @nb_lines OUTPUT, @msgid OUTPUT, @reutrn OUTPUT) at [" + db.getOracleServer + "]; " _
            & "SELECT @message as MSG, @nb_lines as NUM_LINES, @msgid as MSGID, @reutrn as REUTRN; END;"


'Debug.Print loadPurchaseConditionsDataToInterfaceTable

End Function

Function selectPurchaseConditionsDataFromInterfaceTable(msgid As String, barcodes As String) As String
    
    stupci = "" _
& " TNUMSGID, TNULNLIG, TNUCNUF, TNUSUPDESC, TNUCCOM, TNUAGRP, TNUCEXR, TNUADESC, TNULV, TNULU, TNUSITE, TNUSDESC, TNUPACH, TNUPASTPACH, TNUFUTPACH, TNUUAPP, TNUNNC, TNUEXNNC, " _
& " to_char(TNUPADDEB, ''dd.mm.yyyy''), to_char(TNUPADFIN, ''dd.mm.yyyy''), TNUTCP, TNUVAL601, TNUUAPP601, to_char(TNUDDEB601, ''dd.mm.yyyy''), to_char(TNUDFIN601, ''dd.mm.yyyy''), " _
& " TNUPAST601, TNUFUT601, TNUVAL602, TNUUAPP602, to_char(TNUDDEB602, ''dd.mm.yyyy''), to_char(TNUDFIN602, ''dd.mm.yyyy''), TNUPAST602, TNUFUT602, TNUVAL603, TNUUAPP603, " _
& " to_char(TNUDDEB603, ''dd.mm.yyyy''), to_char(TNUDFIN603, ''dd.mm.yyyy''), TNUPAST603, TNUFUT603, TNUVAL604, TNUUAPP604, to_char(TNUDDEB604, ''dd.mm.yyyy''), " _
& " to_char(TNUDFIN604, ''dd.mm.yyyy''), TNUPAST604, TNUFUT604, TNUVAL605, TNUUAPP605, to_char(TNUDDEB605, ''dd.mm.yyyy''), to_char(TNUDFIN605, ''dd.mm.yyyy''), TNUPAST605, " _
& " TNUFUT605, TNUVAL606, TNUUAPP606, to_char(TNUDDEB606, ''dd.mm.yyyy''), to_char(TNUDFIN606, ''dd.mm.yyyy''), TNUPAST606, TNUFUT606, TNUUSER, TNUSTRT, TNUSERRID, TNUSDTRT, " _
& " TNUSMESS , TNUGTRT, TNUGERRID, TNUGDTRT, TNUGMESS, TNUDCRE, TNUDMAJ, TNUTIL, TNUNMAJ "

    tmp = "(select arccode from artvl, artul, artuv, artcoca where ARUSEQVL = ARLSEQVL and arlcexr = TNUCEXR and pkartstock.RecupCinlUVC(123, arucinl) = arvcinv and arutypul = 1 " _
    & "and arccinv = arvcinv and trunc(current_date) between arcddeb and arcdfin and arcieti = 1 and rownum = 1)"

    selectPurchaseConditionsDataFromInterfaceTable = "EXEC ('" _
            & "SELECT " & stupci & ", (select arccode from artvl, artul, artuv, artcoca where ARUSEQVL = ARLSEQVL and ARLCEXVL = TNULV and arlcexr = TNUCEXR and pkartstock.RecupCinlUVC(123, arucinl) = arvcinv and arutypul = 1 and arccinv = arvcinv and trunc(current_date) between arcddeb and arcdfin and arcieti = 1 and rownum = 1) as BARKOD, " _
            & "(SELECT pkattrival.getLibelleCourtAttribut(123, aatccla, aatcatt, ''GB'') FROM artattri, artrac where aatcinr = artcinr and artcexr = TNUCEXR and aatccla = ''PRINCIP'' and trunc(current_date) between aatddeb and aatdfin ) PRINCIPAL, " _
            & "(SELECT max(uatcatt) FROM artuvattri, artuv where uatcinv = arvcinv and arvcexr = TNUCEXR and uatccla = ''ASORT'' and trunc(current_date) between uatddeb and uatdfin ) ASORTIMAN " _
            & "FROM SIS14_NAB_UVJETI " _
            & "WHERE TNUMSGID=''" & msgid & "'' " _
            & "AND (''-1'' in (" & barcodes & ") OR " & tmp & " in (" & barcodes & ")) " _
            & "ORDER BY TNULNLIG') at [" + db.getOracleServer + "];"
End Function

Function updatePurchaseCondition(TNUPACH As Variant, TNUUAPP As String, tnunnc As String, TNUEXNNC As String, TNUPADDEB As Date, TNUPADFIN As Date, TNUTCP As Variant, TNUVAL601 As Variant, TNUUAPP601 As String, TNUDDEB601 As Date, TNUDFIN601 As Date, TNUVAL602 As Variant, TNUUAPP602 As String, TNUDDEB602 As Date, TNUDFIN602 As Date, TNUVAL603 As Variant, TNUUAPP603 As String, TNUDDEB603 As Date, TNUDFIN603 As Date, TNUVAL604 As Variant, TNUUAPP604 As String, TNUDDEB604 As Date, TNUDFIN604 As Date, TNUVAL605 As Variant, TNUUAPP605 As String, TNUDDEB605 As Date, TNUDFIN605 As Date, TNUVAL606 As Variant, TNUUAPP606 As String, TNUDDEB606 As Date, TNUDFIN606 As Date, TNUMSGID As String, TNULNLIG As String) As String
    updatePurchaseCondition = "EXEC ('UPDATE SIS14_NAB_UVJETI " _
        & "SET TNUUSER = ''" & utils.getUserName & "'', " _
        & "TNUPACH = " & utils.getPriceValue(TNUPACH) & ", TNUUAPP = ''" & TNUUAPP & "'', TNUNNC = ''" & tnunnc & "'', TNUEXNNC = ''" & TNUEXNNC & "'', TNUPADDEB = " & utils.getDateString(TNUPADDEB) & ", TNUPADFIN = " & utils.getDateString(TNUPADFIN) & ", TNUTCP = " & utils.getPriceValue(TNUTCP) & ","
    
    If Not IsEmpty(TNUVAL601) Then
        updatePurchaseCondition = updatePurchaseCondition & " TNUVAL601 = " & utils.getPriceValue(TNUVAL601) & ", TNUUAPP601 = ''" & TNUUAPP601 & "'', TNUDDEB601 = " & utils.getDateString(TNUDDEB601) & ", TNUDFIN601 = " & utils.getDateString(TNUDFIN601) & ","
    End If
    If Not IsEmpty(TNUVAL602) Then
        updatePurchaseCondition = updatePurchaseCondition & " TNUVAL602 = " & utils.getPriceValue(TNUVAL602) & ", TNUUAPP602 = ''" & TNUUAPP602 & "'', TNUDDEB602 = " & utils.getDateString(TNUDDEB602) & ", TNUDFIN602 = " & utils.getDateString(TNUDFIN602) & ","
    End If
    If Not IsEmpty(TNUVAL603) Then
        updatePurchaseCondition = updatePurchaseCondition & " TNUVAL603 = " & utils.getPriceValue(TNUVAL603) & ", TNUUAPP603 = ''" & TNUUAPP603 & "'', TNUDDEB603 = " & utils.getDateString(TNUDDEB603) & ", TNUDFIN603 = " & utils.getDateString(TNUDFIN603) & ","
    End If
    If Not IsEmpty(TNUVAL604) Then
        updatePurchaseCondition = updatePurchaseCondition & " TNUVAL604 = " & utils.getPriceValue(TNUVAL604) & ", TNUUAPP604 = ''" & TNUUAPP604 & "'', TNUDDEB604 = " & utils.getDateString(TNUDDEB604) & ", TNUDFIN604 = " & utils.getDateString(TNUDFIN604) & ","
    End If
    If Not IsEmpty(TNUVAL605) Then
        updatePurchaseCondition = updatePurchaseCondition & " TNUVAL605 = " & utils.getPriceValue(TNUVAL605) & ", TNUUAPP605 = ''" & TNUUAPP605 & "'', TNUDDEB605 = " & utils.getDateString(TNUDDEB605) & ", TNUDFIN605 = " & utils.getDateString(TNUDFIN605) & ","
    End If
    If Not IsEmpty(TNUVAL606) Then
        updatePurchaseCondition = updatePurchaseCondition & " TNUVAL606 = " & utils.getPriceValue(TNUVAL606) & ", TNUUAPP606 = ''" & TNUUAPP606 & "'', TNUDDEB606 = " & utils.getDateString(TNUDDEB606) & ", TNUDFIN606 = " & utils.getDateString(TNUDFIN606) & ","
    End If
    
    updatePurchaseCondition = updatePurchaseCondition & " TNUSTRT = 0, TNUDMAJ = sysdate, TNUNMAJ = TNUNMAJ + 1 WHERE TNUMSGID = " & TNUMSGID & " AND TNULNLIG = " & TNULNLIG & "') at [" + db.getOracleServer + "];"

End Function


Function insertASISTATUS(msgid As String) As String
    insertASISTATUS = "EXEC ('" _
    & "INSERT INTO asi_status (SASMSGID, SASAPPLI, SASSSTAT, SASSREAD, SASSPROC, SASSERR, SASGSTAT, SASGREAD, SASGPROC, SASGERR, SASDCRE, SASDMAJ, SASNMAJ, SASUTIL) " _
    & "VALUES (''" + msgid + "'', ''SIS_14'', 0, 0, 0, 0, 0, 0, 0, 0, sysdate, sysdate, 0, ''gb'')') at [" + db.getOracleServer + "];"

End Function


