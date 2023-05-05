Attribute VB_Name = "queries"
Function selectMSGID() As String
    selectMSGID = "EXEC ('select asi_seq_msgid.nextval from dual') at [" + db.getOracleServer + "];"
End Function

Function selectSEQ() As String
    selectSEQ = "EXEC ('select seq_intcdenseq.nextval from dual') at [" + db.getOracleServer + "];"
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


Function loadOrders(site As String, deliveryDate As String, barcodes As String, stores As String)

    loadOrders = "EXEC [Excel].[GetGoldRasterOrders_prod] " _
       & "@site = N'" + site + "', " _
       & "@deliveryDate = N'" + deliveryDate + "', " _
       & "@barcodes = N'" + barcodes + "', " _
       & "@stores = N'" + stores + "'"

End Function


Function checkNonProccesed()

    checkNonProccesed = "EXEC ('"
    checkNonProccesed = checkNonProccesed + "SELECT * FROM INTCDE WHERE INTSTAT = 0 "
    checkNonProccesed = checkNonProccesed + "AND INTUTIL = ''" + CStr(Left(utils.getUserName, 12)) + "''"
    checkNonProccesed = checkNonProccesed + "') at [" + db.getOracleServer + "]; "

End Function

Function deleteNonProccesed()

    deleteNonProccesed = "EXEC ('"
    deleteNonProccesed = deleteNonProccesed + "DELETE FROM INTCDE WHERE INTSTAT = 0 "
    deleteNonProccesed = deleteNonProccesed + "AND INTUTIL = ''" + CStr(Left(utils.getUserName, 12)) + "''"
    deleteNonProccesed = deleteNonProccesed + "') at [" + db.getOracleServer + "]; "

End Function


Function loadDBResponse(msgid As String)
    loadDBResponse = "EXEC ('SELECT pkartcoca.get_closestEAN(123,arvcinv) EAN, pkstrucobj.get_desc(123,arvcinr,''HR'') NAZIV, INTSITE, INTQTEC, to_char(INTDCOM, ''dd-MM-yyyy''), to_char(INTDLIV, ''dd-MM-yyyy hh:mi''), "
    loadDBResponse = loadDBResponse + "INTID, INTLCDE, INTCNUF, INTCCOM, INTNFILF, INTFILC, INTCONF, INTGREL, INTDEVI,INTCOUC, INTTXCH, INTCOM1, INTCOM2, INTENLEV, "
    loadDBResponse = loadDBResponse + "INTDLIM, INTCODE, INTRCOM, INTCEXVA, INTCEXVL, INTUAUVC, INTNEGO, INTORDR, INTSTAT, INTCEXGLO, INTNOOE, INTFLUX, INTFSTA, "
    loadDBResponse = loadDBResponse + "INTLDIST, INTLDNO, INTETAT, INTSITLI, INTPACH, INTCOML1, INTCLCUS, INTURG, INTEXT, INTESCO, INTNJESC, INTPORI, INTINCO, "
    loadDBResponse = loadDBResponse + "INTLIEU2, INTTRSP, INTFRAN, INTVOLI, INTPDSI, INTTYIM, INTDBAS, INTDDEP, INTCRED, INTJOUR, INTDARR, INTMREG, INTDDS, INTNBJM, "
    loadDBResponse = loadDBResponse + "INTDVAL, INTDPAI, INTNSEQ, INTNLIG, INTNLEN, INTFICH, INTCACT, INTNERR, INTMESS, to_char(INTDCRE, ''YYYY-MM-DD-HH24-MI-SS''), to_char(INTDMAJ, ''YYYY-MM-DD-HH24-MI-SS''), INTUTIL, to_char(INTDTRT, ''dd-MM-yyyy''), INTCTVA, "
    loadDBResponse = loadDBResponse + "INTUAPP, INTALTF, INTTYPUL, INTCEXOGL, INTCEXOPS, INTNROUTE, INTLIEU, INTVALOF, INTMOTIF, INTTEL, INTORI, INTCSIN, INTCTLA, "
    loadDBResponse = loadDBResponse + "INTIRECYC, INTCRGP, INTFLIR, INTNOLV, INTDRAM, INTPVSA, INTPVSR, INTPRFA, INTMTDR, INTMTVI, INTGRA, INTDENVREC, INTCEAN, INTCEXTJF, "
    loadDBResponse = loadDBResponse + "INTEDOU, INTRDOU, INTDENLEV, INTREFEXT, INTCTRL, INTFVSA, INTFVSR, INTCODLOG, INTCODCAI, INTUEREMP, INTCINB, INTNOLIGN, INTPROPER "
    loadDBResponse = loadDBResponse + "FROM INTCDE, ARTUV WHERE ARVCEXR = INTCODE "
    loadDBResponse = loadDBResponse + "AND INTUTIL = ''" + CStr(Left(utils.getUserName, 12)) + "''"
    loadDBResponse = loadDBResponse + "AND INTFICH = ''" + CStr(msgid) + "''"
    loadDBResponse = loadDBResponse + "') at [" + db.getOracleServer + "];"
    
    'to_char(current_date, ''YYYY-MM-DD-HH24-MI-SS'')
    
End Function

Function insertOrder(row As Long, msgid As String, seq As String)

    insertOrder = "EXEC('INSERT INTO INTCDE (INTID, INTLCDE, INTSITE, INTCNUF, INTCCOM, INTNFILF, INTFILC, INTCONF, INTGREL, INTDEVI, INTCOUC, INTTXCH, INTCOM1, INTCOM2, INTENLEV, INTDCOM, "
    insertOrder = insertOrder + "INTDLIV, INTDLIM, INTCODE, INTRCOM, INTCEXVA, INTCEXVL, INTQTEC, INTUAUVC, INTNEGO, INTORDR, INTSTAT, INTCEXGLO, INTNOOE, INTFLUX, INTFSTA, INTLDIST, INTLDNO, INTETAT, "
    insertOrder = insertOrder + "INTSITLI, INTPACH, INTCOML1, INTCLCUS, INTURG, INTEXT, INTESCO, INTNJESC, INTPORI, INTINCO, INTLIEU2, INTTRSP, INTFRAN, INTVOLI, INTPDSI, INTTYIM, INTDBAS, INTDDEP, INTCRED, "
    insertOrder = insertOrder + "INTJOUR, INTDARR, INTMREG, INTDDS, INTNBJM, INTDVAL, INTDPAI, INTNSEQ, INTNLIG, INTNLEN, INTFICH, INTCACT, INTNERR, INTMESS, INTDCRE, INTDMAJ, INTUTIL, INTDTRT, INTCTVA, INTUAPP, "
    insertOrder = insertOrder + "INTALTF, INTTYPUL, INTCEXOGL, INTCEXOPS, INTNROUTE, INTLIEU, INTVALOF, INTMOTIF, INTTEL, INTORI, INTCSIN, INTCTLA, INTIRECYC, INTCRGP, INTFLIR, INTNOLV, INTDRAM, INTPVSA, INTPVSR, "
    insertOrder = insertOrder + "INTPRFA, INTMTDR, INTMTVI, INTGRA, INTDENVREC, INTCEAN, INTCEXTJF, INTEDOU, INTRDOU, INTDENLEV, INTREFEXT, INTCTRL, INTFVSA, INTFVSR, INTCODLOG, INTCODCAI, INTUEREMP, INTCINB, INTNOLIGN, "
    insertOrder = insertOrder + "INTPROPER) VALUES ("
        
        insertOrder = insertOrder + "''-1'', " 'INTID
        insertOrder = insertOrder + "NULL, " 'INTLCDE
        insertOrder = insertOrder + CStr(Range(cfg.getcolINTSITE & row).Value) + ", " 'INTSITE
        insertOrder = insertOrder + CStr(Range(cfg.getcolINTCNUF & row).Value) + ", " 'INTCNUF
        insertOrder = insertOrder + "''" + CStr(Range(cfg.getcolINTCCOM & row).Value) + "'', " 'INTCCOM
        insertOrder = insertOrder + CStr(Range(cfg.getcolINTNFILF & row).Value) + ", " 'INTNFILF
        insertOrder = insertOrder + "1, " 'INTFILC
        insertOrder = insertOrder + "0, " 'INTCONF
        insertOrder = insertOrder + "1, " 'INTGREL
        insertOrder = insertOrder + "NULL, " 'INTDEVI
        insertOrder = insertOrder + "0, " 'INTCOUC
        insertOrder = insertOrder + "NULL, " 'INTTXCH
        insertOrder = insertOrder + "''RASTER_" + Format(Date, "yyyy-mm-dd") + "'', " 'INTCOM1
        insertOrder = insertOrder + "NULL, " 'INTCOM2
        insertOrder = insertOrder + "0, " 'INTENLEV
        insertOrder = insertOrder + "trunc(sysdate), " 'INTDCOM
        insertOrder = insertOrder + "to_date(''" + CStr(Range(cfg.getcolINTDLIV & row).Value) + "'',''dd-mm-yyyy hh24:mi''), " 'INTDLIV
        insertOrder = insertOrder + "NULL, " 'INTDLIM
        insertOrder = insertOrder + CStr(Range(cfg.getcolINTCODE & row).Value) + ", " 'INTCODE
        insertOrder = insertOrder + "''-1'', " 'INTRCOM
        insertOrder = insertOrder + CStr(Range(cfg.getcolINTCEXVA & row).Value) + ", " 'INTCEXVA
        insertOrder = insertOrder + CStr(Range(cfg.getcolINTCEXVL & row).Value) + ", " 'INTCEXVL
        If (Len(Range(cfg.getcolINTQTEC & row).Value) > 0) Then
            insertOrder = insertOrder + CStr(Range(cfg.getcolINTQTEC & row).Value) + ", " 'INTQTEC
        Else
            insertOrder = insertOrder + "NULL, " 'INTQTEC
        End If
        insertOrder = insertOrder + "NULL, " 'INTUAUVC
        insertOrder = insertOrder + "NULL, " 'INTNEGO
        insertOrder = insertOrder + "NULL, " 'INTORDR
        insertOrder = insertOrder + "0, " 'INTSTAT
        insertOrder = insertOrder + "NULL, " 'INTCEXGLO
        insertOrder = insertOrder + "NULL, " 'INTNOOE
        insertOrder = insertOrder + "1, " 'INTFLUX
        insertOrder = insertOrder + "NULL, " 'INTFSTA
        insertOrder = insertOrder + "0, " 'INTLDIST
        insertOrder = insertOrder + "NULL, " 'INTLDNO
        insertOrder = insertOrder + "5, " 'INTETAT
        
        If Len(Range(cfg.getcolINTSITLI & row).Value) = 5 Then
            'interno skladište - gledamo po duljni šifre skladišta/dobavljaèa
            insertOrder = insertOrder + CStr(Range(cfg.getcolINTSITLI & row).Value) + ", " 'INTSITLI
        Else
            'vanjski dobavljaè
            insertOrder = insertOrder + "NULL, " 'INTSITLI
        End If
        
        insertOrder = insertOrder + "NULL, " 'INTPACH
        insertOrder = insertOrder + "NULL, " 'INTCOML1
        insertOrder = insertOrder + "NULL, " 'INTCLCUS
        insertOrder = insertOrder + "0, " 'INTURG
        insertOrder = insertOrder + "NULL, " 'INTEXT
        insertOrder = insertOrder + "NULL, " 'INTESCO
        insertOrder = insertOrder + "NULL, " 'INTNJESC
        insertOrder = insertOrder + "NULL, " 'INTPORI
        insertOrder = insertOrder + "NULL, " 'INTINCO
        insertOrder = insertOrder + "NULL, " 'INTLIEU2
        insertOrder = insertOrder + "NULL, " 'INTTRSP
        insertOrder = insertOrder + "0, " 'INTFRAN
        insertOrder = insertOrder + "NULL, " 'INTVOLI
        insertOrder = insertOrder + "NULL, " 'INTPDSI
        insertOrder = insertOrder + "NULL, " 'INTTYIM
        insertOrder = insertOrder + "NULL, " 'INTDBAS
        insertOrder = insertOrder + "NULL, " 'INTDDEP
        insertOrder = insertOrder + "NULL, " 'INTCRED
        insertOrder = insertOrder + "NULL, " 'INTJOUR
        insertOrder = insertOrder + "NULL, " 'INTDARR
        insertOrder = insertOrder + "NULL, " 'INTMREG
        insertOrder = insertOrder + "NULL, " 'INTDDS
        insertOrder = insertOrder + "NULL, " 'INTNBJM
        insertOrder = insertOrder + "NULL, " 'INTDVAL
        insertOrder = insertOrder + "NULL, " 'INTDPAI
        insertOrder = insertOrder + CStr(seq) + ", " 'INTNSEQ
        insertOrder = insertOrder + "-1, " 'INTNLIG
        insertOrder = insertOrder + "NULL, " 'INTNLEN
        insertOrder = insertOrder + "''" + CStr(msgid) + "'', " 'INTFICH
        insertOrder = insertOrder + "1, " 'INTCACT
        insertOrder = insertOrder + "NULL, " 'INTNERR
        insertOrder = insertOrder + "NULL, " 'INTMESS
        insertOrder = insertOrder + "CURRENT_DATE, " 'INTDCRE
        insertOrder = insertOrder + "CURRENT_DATE, " 'INTDMAJ
        insertOrder = insertOrder + "''" + CStr(Left(utils.getUserName, 12)) + "'', " 'INTUTIL
        insertOrder = insertOrder + "trunc(sysdate), " 'INTDTRT
        insertOrder = insertOrder + "NULL, " 'INTCTVA
        insertOrder = insertOrder + "NULL, " 'INTUAPP
        insertOrder = insertOrder + "0, " 'INTALTF
        insertOrder = insertOrder + CStr(Range(cfg.getcolINTTYPUL & row).Value) + ", " 'INTTYPUL
        insertOrder = insertOrder + "NULL, " 'INTCEXOGL
        insertOrder = insertOrder + "NULL, " 'INTCEXOPS
        insertOrder = insertOrder + "NULL, " 'INTNROUTE
        insertOrder = insertOrder + "NULL, " 'INTLIEU
        insertOrder = insertOrder + "NULL, " 'INTVALOF
        insertOrder = insertOrder + "NULL, " 'INTMOTIF
        insertOrder = insertOrder + "NULL, " 'INTTEL
        insertOrder = insertOrder + "906, " 'INTORI
        insertOrder = insertOrder + "NULL, " 'INTCSIN
        insertOrder = insertOrder + "1, " 'INTCTLA
        insertOrder = insertOrder + "0, " 'INTIRECYC
        insertOrder = insertOrder + "NULL, " 'INTCRGP
        insertOrder = insertOrder + "0, " 'INTFLIR
        insertOrder = insertOrder + "NULL, " 'INTNOLV
        insertOrder = insertOrder + "NULL, " 'INTDRAM
        insertOrder = insertOrder + "NULL, " 'INTPVSA
        insertOrder = insertOrder + "NULL, " 'INTPVSR
        insertOrder = insertOrder + "NULL, " 'INTPRFA
        insertOrder = insertOrder + "NULL, " 'INTMTDR
        insertOrder = insertOrder + "NULL, " 'INTMTVI
        insertOrder = insertOrder + "NULL, " 'INTGRA
        insertOrder = insertOrder + "NULL, " 'INTDENVREC
        insertOrder = insertOrder + "NULL, " 'INTCEAN
        insertOrder = insertOrder + "NULL, " 'INTCEXTJF
        insertOrder = insertOrder + "NULL, " 'INTEDOU
        insertOrder = insertOrder + "NULL, " 'INTRDOU
        insertOrder = insertOrder + "NULL, " 'INTDENLEV
        insertOrder = insertOrder + "NULL, " 'INTREFEXT
        insertOrder = insertOrder + "NULL, " 'INTCTRL
        insertOrder = insertOrder + "NULL, " 'INTFVSA
        insertOrder = insertOrder + "NULL, " 'INTFVSR
        insertOrder = insertOrder + "''-1'', " 'INTCODLOG
        insertOrder = insertOrder + "''-1'', " 'INTCODCAI
        insertOrder = insertOrder + "NULL, " 'INTUEREMP
        insertOrder = insertOrder + "NULL, " 'INTCINB
        insertOrder = insertOrder + "NULL, " 'INTNOLIGN
        insertOrder = insertOrder + "NULL" 'INTPROPER
    
    insertOrder = insertOrder + ")') at [" + db.getOracleServer + "]; "

End Function

Function getDocumentVersion(doc_name As String) As String
     getDocumentVersion = "SELECT TOP 1 [document_version] FROM [excel].[excel_document_versions] WHERE [document_name] = '" & doc_name & "'"
     getDocumentVersion = getDocumentVersion + " ORDER BY [timestamp] DESC"
End Function
