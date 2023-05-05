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

Function selectReceptions(site As String, receptions As String, deliveryNums) As String

    selectReceptions = "EXEC ('SELECT sernusr BROJ_PRIJEMA, serbliv BROJ_DOSTAVNICE"
    selectReceptions = selectReceptions & " ,pkcodeartrecept.getCodeArticleCde(1, nvl(nvl(sdrcode, dcdcode), artcexr), artcinr, arucinl, aruseqvl, sercfin, serccin, "
    selectReceptions = selectReceptions & " serfilf, sdrsite, to_char(sdrsdrc, ''DD/MM/RR''), ''GB'', dcdcexr, dcdcoca, dcdcoul, dcdcean, dcdrefc) KOD_ARTIKLA, "
    selectReceptions = selectReceptions & " sdrcexr CEXR, "
    selectReceptions = selectReceptions & " sdrrefc REFC, "
    selectReceptions = selectReceptions & " SDRCODE EAN, "
    selectReceptions = selectReceptions & " (select count(1) from storemre a where a.srrcinrec=sdrcinrec and a.srrnlrec=sdrnlrec and SRRTREM=904) PN_FLAG, "
    selectReceptions = selectReceptions & " sdrcexvl LV, "
    selectReceptions = selectReceptions & " pkparpostes.get_postlibl(1, 10, 731, NVL(arutypul, 0), ''HR'') OPIS_LV, "
    selectReceptions = selectReceptions & " pkstrucobj.get_desc(1, aruclibl, ''HR'') NAZIV_ARTIKLA, "
    selectReceptions = selectReceptions & " SDRCTVA PDV_GRUPA, "
    selectReceptions = selectReceptions & " pkparpostes.get_postlibl(123,0,7,SDRCTVA,''HR'') NAZIV_PDV_GRUPE, "
    selectReceptions = selectReceptions & " SDRTTVA PDV_STOPA, "
    selectReceptions = selectReceptions & " decode( artustk, 1, SDRQTES,SDRPDSS) KOLICINA, "
    selectReceptions = selectReceptions & " DECODE(PKCTRLACCES.ctrl_acces_saisie(1, 10, ''gb'', 13), 1, sdrpnfa, null) NC_EUR, "
    selectReceptions = selectReceptions & " pkparpostes.get_postlibl(123,0,805,serdevi,''HR'') OPIS_NC_EUR, "
    selectReceptions = selectReceptions & " case when serdevi=978 then"
    selectReceptions = selectReceptions & "     round(DECODE(PKCTRLACCES.ctrl_acces_saisie(1, 10, ''gb'', 13), 1, sdrpnfa, null) * 7.5345 ,2) "
    selectReceptions = selectReceptions & " Else "
    selectReceptions = selectReceptions & "     DECODE(PKCTRLACCES.ctrl_acces_saisie(1, 10, ''gb'', 13), 1, sdrpnfa, null) "
    selectReceptions = selectReceptions & " end NC_HRK, "
    selectReceptions = selectReceptions & " pkparpostes.get_postlibl(123,0,805,191,''HR'') OPIS_NC_HRK, "
    selectReceptions = selectReceptions & " serccin CCIN,"
    selectReceptions = selectReceptions & " pkfouccom.get_NumContrat(123,serccin) CCOM,"
    selectReceptions = selectReceptions & " serfilf filf"
    
    selectReceptions = selectReceptions & "  FROM artrac, stodetre, cdedetcde, artul, stoentre "
    
    selectReceptions = selectReceptions & " where sdrsite = " & site
    selectReceptions = selectReceptions & "  AND sdrcinrec in( select sercinrec from stoentre where sersite = sdrsite and (sernusr in ( "
    selectReceptions = selectReceptions & receptions & ") or serbliv in ("
    selectReceptions = selectReceptions & deliveryNums & ")))"
    selectReceptions = selectReceptions & " AND artcinr = sdrcinr AND dcdcincde(+) = sercincde AND dcdsite(+) = sersite AND DCDNOLIGN(+) = SDRNLCDE "
    selectReceptions = selectReceptions & " AND nvl(DCDNOPS(+), 0) = nvl(SDRNOPS, 0) AND arucinr = sdrcinr AND arucinl = sdrcinla AND sdrcinrec = sercinrec "
    selectReceptions = selectReceptions & " AND dcdnligp(+) = sdrnligp ORDER BY nvl(sdrnlrecorig, sdrnlrec), sdrnlrec, sdrnligp "
    
    selectReceptions = selectReceptions & "') at [" + db.getOracleServer + "];"

End Function


Function insertInvoiceHeader(sercfin As String, serccin As String, supcnuf As String, supccom As String, invoiceNum As String, invoiceDate As String, payDate As String, _
    rowNum As Integer, serfilf As String, util As String, totalAmmount As Double, fich As String) As String
    'intcfinv
    insertInvoiceHeader = "EXEC(' "
    
    insertInvoiceHeader = insertInvoiceHeader & "INSERT INTO intcfinv (CFICFIN ,CFICCIN, CFICFEX, CFICCEX, CFISITE, CFIINVID, CFIDATE, CFITYRP, CFITYPE, CFIDEV, CFIDECH, CFIDECHDEM, CFIMORP, CFIDKRO, "
    'insertInvoiceHeader = insertInvoiceHeader & " CFIIDEN, "
    insertInvoiceHeader = insertInvoiceHeader & "  CFIDVAL, CFIMFAC, CFITOL, CFITYTRT, CFITYPTOL, CFIDFLGI, CFISTAT, CFIDTRT, CFIFICH, CFINLIG, CFIDCRE, CFIDMAJ, CFIUTIL, CFINFILF, "
    insertInvoiceHeader = insertInvoiceHeader & "  CFICNPAY, CFIEMB, CFIREMA, CFIREMP, CFICASSAI, CFIMOPAYREF, CFIPAYREF, CFIUTILSAISI, CFIUTILMAJ) "
    insertInvoiceHeader = insertInvoiceHeader & "VALUES ( "
    insertInvoiceHeader = insertInvoiceHeader & sercfin
    insertInvoiceHeader = insertInvoiceHeader & ", " & serccin
    insertInvoiceHeader = insertInvoiceHeader & ", ''" & supcnuf & "'' "
    insertInvoiceHeader = insertInvoiceHeader & ", ''" & supccom & "'' "
    insertInvoiceHeader = insertInvoiceHeader & ", 21210 "
    insertInvoiceHeader = insertInvoiceHeader & ", ''" & invoiceNum & "'' "
    insertInvoiceHeader = insertInvoiceHeader & ", " & invoiceDate
    insertInvoiceHeader = insertInvoiceHeader & ", 1 "
    insertInvoiceHeader = insertInvoiceHeader & ", 1 "
    insertInvoiceHeader = insertInvoiceHeader & ", 978 "
    insertInvoiceHeader = insertInvoiceHeader & ", " & payDate
    insertInvoiceHeader = insertInvoiceHeader & ", " & payDate
    insertInvoiceHeader = insertInvoiceHeader & " , 1 "
    insertInvoiceHeader = insertInvoiceHeader & ", " & invoiceDate
    'insertInvoiceHeader = insertInvoiceHeader &    "   --,prazno(ako bude trebalo dohvatiom OIB kansije) "
    insertInvoiceHeader = insertInvoiceHeader & ", " & invoiceDate
    insertInvoiceHeader = insertInvoiceHeader & ", " & Replace(CStr(totalAmmount), ",", ".") 'total fakture prepisano sa raèuna (u eurima iznos)
    insertInvoiceHeader = insertInvoiceHeader & ", 1 "
    insertInvoiceHeader = insertInvoiceHeader & ", 2 "
    insertInvoiceHeader = insertInvoiceHeader & ", 0 "
    insertInvoiceHeader = insertInvoiceHeader & ", 0 "
    insertInvoiceHeader = insertInvoiceHeader & ", 0 "
    insertInvoiceHeader = insertInvoiceHeader & ", sysdate "
    insertInvoiceHeader = insertInvoiceHeader & " , ''" & fich & "''"
    insertInvoiceHeader = insertInvoiceHeader & ", " & rowNum
    insertInvoiceHeader = insertInvoiceHeader & ", sysdate "
    insertInvoiceHeader = insertInvoiceHeader & ", sysdate "
    insertInvoiceHeader = insertInvoiceHeader & " , ''" & Left(util, 12) & "''"
    insertInvoiceHeader = insertInvoiceHeader & ", " & serfilf
    insertInvoiceHeader = insertInvoiceHeader & " , 0 "
    insertInvoiceHeader = insertInvoiceHeader & " , 0 "
    insertInvoiceHeader = insertInvoiceHeader & " , 1 "
    insertInvoiceHeader = insertInvoiceHeader & " , 0 "
    insertInvoiceHeader = insertInvoiceHeader & " , 0 "
    insertInvoiceHeader = insertInvoiceHeader & " , 1 "
    insertInvoiceHeader = insertInvoiceHeader & " , ''01'' "
    insertInvoiceHeader = insertInvoiceHeader & " , ''" & Left(util, 12) & "''"
    insertInvoiceHeader = insertInvoiceHeader & " , ''" & Left(util, 12) & "''"
    insertInvoiceHeader = insertInvoiceHeader & ") "
    
    insertInvoiceHeader = insertInvoiceHeader & "') at [" + db.getOracleServer + "];"
End Function


Function insertVatRates(supcnuf As String, invoiceNum As String, deliveryNum As String, vatRate As Double, netAmmount As Double, vatAmmount As Double, _
    site As String, fich As String, util As String, rowNum As Long) As String
    'intcfbl
    insertVatRates = "EXEC(' "
    
    insertVatRates = insertVatRates & " INSERT INTO intcfbl ("
    insertVatRates = insertVatRates & "     CFBCFEX, CFBINVID, CFBBLID, CFBTYPE, CFBTXTVA, CFBSITE, CFBMONT, CFBTXMNT, CFBDFLGI, "
    insertVatRates = insertVatRates & "     CFBSTAT, CFBDTRT, CFBFICH, CFBNLIG, CFBDCRE, CFBDMAJ, CFBUTIL)  "
    insertVatRates = insertVatRates & " VALUES("
    insertVatRates = insertVatRates & "''" & supcnuf & "'' "
    insertVatRates = insertVatRates & ", ''" & invoiceNum & "'' "
    insertVatRates = insertVatRates & ", ''" & deliveryNum & "'' " 'broj dostavnice serbliv
    insertVatRates = insertVatRates & ", 1"
    insertVatRates = insertVatRates & ", " & CStr(vatRate) 'stopa poreza sa fakture (u eurima iznos)
    insertVatRates = insertVatRates & ", " & site
    insertVatRates = insertVatRates & ", " & Replace(CStr(netAmmount), ",", ".") 'osnovica za porez sa fakture (u eurima iznos)
    insertVatRates = insertVatRates & ", " & Replace(CStr(vatAmmount), ",", ".") 'iznos poreza sa fakture (u eurima iznos)
    insertVatRates = insertVatRates & ",1 "
    insertVatRates = insertVatRates & ",0 "
    insertVatRates = insertVatRates & ", sysdate "
    insertVatRates = insertVatRates & " , ''" & fich & "''"
    insertVatRates = insertVatRates & ", " & rowNum
    insertVatRates = insertVatRates & ", sysdate "
    insertVatRates = insertVatRates & ", sysdate "
    insertVatRates = insertVatRates & " , ''" & Left(util, 12) & "''"
    insertVatRates = insertVatRates & ")        "
    
    insertVatRates = insertVatRates & "') at [" + db.getOracleServer + "];"
End Function

Function insertInvoiceLine(supcnuf As String, invoiceNum As String, deliveryNum As String, cexr As String, lv As String, refc As String, xtva As String, _
    qty As Double, nc As Double, site As String, rowNum As Long, fich As String, util As String, ean As String) As String
    'intcfart
    insertInvoiceLine = "EXEC(' "
    
    insertInvoiceLine = insertInvoiceLine & "INSERT INTO intcfart ( "
    insertInvoiceLine = insertInvoiceLine & "CFACFEX, CFAINVID, CFABLID, CFACEXR, CFACEXVL, CFAREFC, CFATXTVA, CFAQTY, CFABTPRX, CFATYPE, CFASITE, "
    insertInvoiceLine = insertInvoiceLine & "CFAQTG, CFADFLGI, CFASTAT, CFADTRT, CFAFICH, CFANLIG, CFADCRE, CFADMAJ, CFAUTIL, CFACODCAI) "
    insertInvoiceLine = insertInvoiceLine & "VALUES( "
    insertInvoiceLine = insertInvoiceLine & "''" & supcnuf & "'' "
    insertInvoiceLine = insertInvoiceLine & ", ''" & invoiceNum & "'' "
    insertInvoiceLine = insertInvoiceLine & ", ''" & deliveryNum & "'' " 'broj dostavnice serbliv
    insertInvoiceLine = insertInvoiceLine & ", ''" & cexr & "'' "
    insertInvoiceLine = insertInvoiceLine & ", " & lv
    insertInvoiceLine = insertInvoiceLine & ", ''" & refc & "'' "
    insertInvoiceLine = insertInvoiceLine & ", " & CStr(xtva)
    insertInvoiceLine = insertInvoiceLine & ", " & Replace(CStr(qty), ",", ".")
    insertInvoiceLine = insertInvoiceLine & ", " & Replace(CStr(nc), ",", ".")
    insertInvoiceLine = insertInvoiceLine & ", 1 "
    insertInvoiceLine = insertInvoiceLine & ", " & site
    insertInvoiceLine = insertInvoiceLine & ", 0 "
    insertInvoiceLine = insertInvoiceLine & ", " & rowNum
    insertInvoiceLine = insertInvoiceLine & ", 0 "
    insertInvoiceLine = insertInvoiceLine & ", sysdate "
    insertInvoiceLine = insertInvoiceLine & " , ''" & fich & "''"
    insertInvoiceLine = insertInvoiceLine & ", " & rowNum
    insertInvoiceLine = insertInvoiceLine & ", sysdate "
    insertInvoiceLine = insertInvoiceLine & ", sysdate "
    insertInvoiceLine = insertInvoiceLine & " , ''" & Left(util, 12) & "''"
    insertInvoiceLine = insertInvoiceLine & " , ''" & ean & "''"
    insertInvoiceLine = insertInvoiceLine & ")"
    
    insertInvoiceLine = insertInvoiceLine & "') at [" + db.getOracleServer + "];"
End Function
