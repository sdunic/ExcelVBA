Attribute VB_Name = "cfg"
Dim colEAN, colNAZIV, colINTSITE, colINTQTEC, colINTDCOM, colINTDLIV, colINTID, colINTLCDE, colINTCNUF, colINTCCOM, colINTNFILF, colINTFILC, colINTCONF, colINTGREL, colINTDEVI As String
Dim colINTCOUC, colINTTXCH, colINTCOM1, colINTCOM2, colINTENLEV, colINTDLIM, colINTCODE, colINTRCOM, colINTCEXVA, colINTCEXVL, colINTUAUVC, colINTNEGO, colINTORDR, colINTSTAT, colINTCEXGLO As String
Dim colINTNOOE, colINTFLUX, colINTFSTA, colINTLDIST, colINTLDNO, colINTETAT, colINTSITLI, colINTPACH, colINTCOML1, colINTCLCUS, colINTURG, colINTEXT, colINTESCO, colINTNJESC, colINTPORI, colINTINCO As String
Dim colINTLIEU2, colINTTRSP, colINTFRAN, colINTVOLI, colINTPDSI, colINTTYIM, colINTDBAS, colINTDDEP, colINTCRED, colINTJOUR, colINTDARR, colINTMREG, colINTDDS, colINTNBJM, colINTDVAL, colINTDPAI, colINTNSEQ As String
Dim colINTNLIG, colINTNLEN, colINTFICH, colINTCACT, colINTNERR, colINTMESS, colINTDCRE, colINTDMAJ, colINTUTIL, colINTDTRT, colINTCTVA, colINTUAPP, colINTALTF, colINTTYPUL, colINTCEXOGL, colINTCEXOPS As String
Dim colINTNROUTE, colINTLIEU, colINTVALOF, colINTMOTIF, colINTTEL, colINTORI, colINTCSIN, colINTCTLA, colINTIRECYC, colINTCRGP, colINTFLIR, colINTNOLV, colINTDRAM, colINTPVSA, colINTPVSR As String
Dim colINTPRFA, colINTMTDR, colINTMTVI, colINTGRA, colINTDENVREC, colINTCEAN, colINTCEXTJF, colINTEDOU, colINTRDOU, colINTDENLEV, colINTREFEXT, colINTCTRL, colINTFVSA, colINTFVSR, colINTCODLOG As String
Dim colINTCODCAI, colINTUEREMP, colINTCINB, colINTNOLIGN, colINTPROPER As String

Dim rsEAN, rsNAZIV, rsINTSITE, rsINTQTEC, rsINTDCOM, rsINTDLIV, rsINTID, rsINTLCDE, rsINTCNUF, rsINTCCOM, rsINTNFILF, rsINTFILC, rsINTCONF, rsINTGREL, rsINTDEVI As Integer
Dim rsINTCOUC, rsINTTXCH, rsINTCOM1, rsINTCOM2, rsINTENLEV, rsINTDLIM, rsINTCODE, rsINTRCOM, rsINTCEXVA, rsINTCEXVL, rsINTUAUVC, rsINTNEGO, rsINTORDR, rsINTSTAT, rsINTCEXGLO As Integer
Dim rsINTNOOE, rsINTFLUX, rsINTFSTA, rsINTLDIST, rsINTLDNO, rsINTETAT, rsINTSITLI, rsINTPACH, rsINTCOML1, rsINTCLCUS, rsINTURG, rsINTEXT, rsINTESCO, rsINTNJESC, rsINTPORI, rsINTINCO As Integer
Dim rsINTLIEU2, rsINTTRSP, rsINTFRAN, rsINTVOLI, rsINTPDSI, rsINTTYIM, rsINTDBAS, rsINTDDEP, rsINTCRED, rsINTJOUR, rsINTDARR, rsINTMREG, rsINTDDS, rsINTNBJM, rsINTDVAL, rsINTDPAI, rsINTNSEQ As Integer
Dim rsINTNLIG, rsINTNLEN, rsINTFICH, rsINTCACT, rsINTNERR, rsINTMESS, rsINTDCRE, rsINTDMAJ, rsINTUTIL, rsINTDTRT, rsINTCTVA, rsINTUAPP, rsINTALTF, rsINTTYPUL, rsINTCEXOGL, rsINTCEXOPS As Integer
Dim rsINTNROUTE, rsINTLIEU, rsINTVALOF, rsINTMOTIF, rsINTTEL, rsINTORI, rsINTCSIN, rsINTCTLA, rsINTIRECYC, rsINTCRGP, rsINTFLIR, rsINTNOLV, rsINTDRAM, rsINTPVSA, rsINTPVSR As Integer
Dim rsINTPRFA, rsINTMTDR, rsINTMTVI, rsINTGRA, rsINTDENVREC, rsINTCEAN, rsINTCEXTJF, rsINTEDOU, rsINTRDOU, rsINTDENLEV, rsINTREFEXT, rsINTCTRL, rsINTFVSA, rsINTFVSR, rsINTCODLOG As Integer
Dim rsINTCODCAI, rsINTUEREMP, rsINTCINB, rsINTNOLIGN, rsINTPROPER As Integer


Sub Init()

    colEAN = "B"
    colNAZIV = "C"
    colINTSITE = "D"
    colINTQTEC = "E"
    colINTDCOM = "F"
    colINTDLIV = "G"
    colINTID = "H"
    colINTLCDE = "I"
    colINTCNUF = "J"
    colINTCCOM = "K"
    colINTNFILF = "L"
    colINTFILC = "M"
    colINTCONF = "N"
    colINTGREL = "O"
    colINTDEVI = "P"
    colINTCOUC = "Q"
    colINTTXCH = "R"
    colINTCOM1 = "S"
    colINTCOM2 = "T"
    colINTENLEV = "U"
    colINTDLIM = "V"
    colINTCODE = "W"
    colINTRCOM = "X"
    colINTCEXVA = "Y"
    colINTCEXVL = "Z"
    colINTUAUVC = "AA"
    colINTNEGO = "AB"
    colINTORDR = "AC"
    colINTSTAT = "AD"
    colINTCEXGLO = "AE"
    colINTNOOE = "AF"
    colINTFLUX = "AG"
    colINTFSTA = "AH"
    colINTLDIST = "AI"
    colINTLDNO = "AJ"
    colINTETAT = "AK"
    colINTSITLI = "AL"
    colINTPACH = "AM"
    colINTCOML1 = "AN"
    colINTCLCUS = "AO"
    colINTURG = "AP"
    colINTEXT = "AQ"
    colINTESCO = "AR"
    colINTNJESC = "AS"
    colINTPORI = "AT"
    colINTINCO = "AU"
    colINTLIEU2 = "AV"
    colINTTRSP = "AW"
    colINTFRAN = "AX"
    colINTVOLI = "AY"
    colINTPDSI = "AZ"
    colINTTYIM = "BA"
    colINTDBAS = "BB"
    colINTDDEP = "BC"
    colINTCRED = "BD"
    colINTJOUR = "BE"
    colINTDARR = "BF"
    colINTMREG = "BG"
    colINTDDS = "BH"
    colINTNBJM = "BI"
    colINTDVAL = "BJ"
    colINTDPAI = "BK"
    colINTNSEQ = "BL"
    colINTNLIG = "BM"
    colINTNLEN = "BN"
    colINTFICH = "BO"
    colINTCACT = "BP"
    colINTNERR = "BQ"
    colINTMESS = "BR"
    colINTDCRE = "BS"
    colINTDMAJ = "BT"
    colINTUTIL = "BU"
    colINTDTRT = "BV"
    colINTCTVA = "BW"
    colINTUAPP = "BX"
    colINTALTF = "BY"
    colINTTYPUL = "BZ"
    colINTCEXOGL = "CA"
    colINTCEXOPS = "CB"
    colINTNROUTE = "CC"
    colINTLIEU = "CD"
    colINTVALOF = "CE"
    colINTMOTIF = "CF"
    colINTTEL = "CG"
    colINTORI = "CH"
    colINTCSIN = "CI"
    colINTCTLA = "CJ"
    colINTIRECYC = "CK"
    colINTCRGP = "CL"
    colINTFLIR = "CM"
    colINTNOLV = "CN"
    colINTDRAM = "CO"
    colINTPVSA = "CP"
    colINTPVSR = "CQ"
    colINTPRFA = "CR"
    colINTMTDR = "CS"
    colINTMTVI = "CT"
    colINTGRA = "CU"
    colINTDENVREC = "CV"
    colINTCEAN = "CW"
    colINTCEXTJF = "CX"
    colINTEDOU = "CY"
    colINTRDOU = "CZ"
    colINTDENLEV = "DA"
    colINTREFEXT = "DB"
    colINTCTRL = "DC"
    colINTFVSA = "DD"
    colINTFVSR = "DE"
    colINTCODLOG = "DF"
    colINTCODCAI = "DG"
    colINTUEREMP = "DH"
    colINTCINB = "DI"
    colINTNOLIGN = "DJ"
    colINTPROPER = "DK"


    rsEAN = 0
    rsNAZIV = 1
    rsINTSITE = 2
    rsINTQTEC = 3
    rsINTDCOM = 4
    rsINTDLIV = 5
    rsINTID = 6
    rsINTLCDE = 7
    rsINTCNUF = 8
    rsINTCCOM = 9
    rsINTNFILF = 10
    rsINTFILC = 11
    rsINTCONF = 12
    rsINTGREL = 13
    rsINTDEVI = 14
    rsINTCOUC = 15
    rsINTTXCH = 16
    rsINTCOM1 = 17
    rsINTCOM2 = 18
    rsINTENLEV = 19
    rsINTDLIM = 20
    rsINTCODE = 21
    rsINTRCOM = 22
    rsINTCEXVA = 23
    rsINTCEXVL = 24
    rsINTUAUVC = 25
    rsINTNEGO = 26
    rsINTORDR = 27
    rsINTSTAT = 28
    rsINTCEXGLO = 29
    rsINTNOOE = 30
    rsINTFLUX = 31
    rsINTFSTA = 32
    rsINTLDIST = 33
    rsINTLDNO = 34
    rsINTETAT = 35
    rsINTSITLI = 36
    rsINTPACH = 37
    rsINTCOML1 = 38
    rsINTCLCUS = 39
    rsINTURG = 40
    rsINTEXT = 41
    rsINTESCO = 42
    rsINTNJESC = 43
    rsINTPORI = 44
    rsINTINCO = 45
    rsINTLIEU2 = 46
    rsINTTRSP = 47
    rsINTFRAN = 48
    rsINTVOLI = 49
    rsINTPDSI = 50
    rsINTTYIM = 51
    rsINTDBAS = 52
    rsINTDDEP = 53
    rsINTCRED = 54
    rsINTJOUR = 55
    rsINTDARR = 56
    rsINTMREG = 57
    rsINTDDS = 58
    rsINTNBJM = 59
    rsINTDVAL = 60
    rsINTDPAI = 61
    rsINTNSEQ = 62
    rsINTNLIG = 63
    rsINTNLEN = 64
    rsINTFICH = 65
    rsINTCACT = 66
    rsINTNERR = 67
    rsINTMESS = 68
    rsINTDCRE = 69
    rsINTDMAJ = 70
    rsINTUTIL = 71
    rsINTDTRT = 72
    rsINTCTVA = 73
    rsINTUAPP = 74
    rsINTALTF = 75
    rsINTTYPUL = 76
    rsINTCEXOGL = 77
    rsINTCEXOPS = 78
    rsINTNROUTE = 79
    rsINTLIEU = 80
    rsINTVALOF = 81
    rsINTMOTIF = 82
    rsINTTEL = 83
    rsINTORI = 84
    rsINTCSIN = 85
    rsINTCTLA = 86
    rsINTIRECYC = 87
    rsINTCRGP = 88
    rsINTFLIR = 89
    rsINTNOLV = 90
    rsINTDRAM = 91
    rsINTPVSA = 92
    rsINTPVSR = 93
    rsINTPRFA = 94
    rsINTMTDR = 95
    rsINTMTVI = 96
    rsINTGRA = 97
    rsINTDENVREC = 98
    rsINTCEAN = 99
    rsINTCEXTJF = 100
    rsINTEDOU = 101
    rsINTRDOU = 102
    rsINTDENLEV = 103
    rsINTREFEXT = 104
    rsINTCTRL = 105
    rsINTFVSA = 106
    rsINTFVSR = 107
    rsINTCODLOG = 108
    rsINTCODCAI = 109
    rsINTUEREMP = 110
    rsINTCINB = 111
    rsINTNOLIGN = 112
    rsINTPROPER = 113



End Sub

Function getcolEAN() As String
    getcolEAN = colEAN
End Function

Function getcolNAZIV() As String
    getcolNAZIV = colNAZIV
End Function

Function getcolINTSITE() As String
    getcolINTSITE = colINTSITE
End Function

Function getcolINTQTEC() As String
    getcolINTQTEC = colINTQTEC
End Function

Function getcolINTDCOM() As String
    getcolINTDCOM = colINTDCOM
End Function

Function getcolINTDLIV() As String
    getcolINTDLIV = colINTDLIV
End Function

Function getcolINTID() As String
    getcolINTID = colINTID
End Function

Function getcolINTLCDE() As String
    getcolINTLCDE = colINTLCDE
End Function

Function getcolINTCNUF() As String
    getcolINTCNUF = colINTCNUF
End Function

Function getcolINTCCOM() As String
    getcolINTCCOM = colINTCCOM
End Function

Function getcolINTNFILF() As String
    getcolINTNFILF = colINTNFILF
End Function

Function getcolINTFILC() As String
    getcolINTFILC = colINTFILC
End Function

Function getcolINTCONF() As String
    getcolINTCONF = colINTCONF
End Function

Function getcolINTGREL() As String
    getcolINTGREL = colINTGREL
End Function

Function getcolINTDEVI() As String
    getcolINTDEVI = colINTDEVI
End Function

Function getcolINTCOUC() As String
    getcolINTCOUC = colINTCOUC
End Function

Function getcolINTTXCH() As String
    getcolINTTXCH = colINTTXCH
End Function

Function getcolINTCOM1() As String
    getcolINTCOM1 = colINTCOM1
End Function

Function getcolINTCOM2() As String
    getcolINTCOM2 = colINTCOM2
End Function

Function getcolINTENLEV() As String
    getcolINTENLEV = colINTENLEV
End Function

Function getcolINTDLIM() As String
    getcolINTDLIM = colINTDLIM
End Function

Function getcolINTCODE() As String
    getcolINTCODE = colINTCODE
End Function

Function getcolINTRCOM() As String
    getcolINTRCOM = colINTRCOM
End Function

Function getcolINTCEXVA() As String
    getcolINTCEXVA = colINTCEXVA
End Function

Function getcolINTCEXVL() As String
    getcolINTCEXVL = colINTCEXVL
End Function

Function getcolINTUAUVC() As String
    getcolINTUAUVC = colINTUAUVC
End Function

Function getcolINTNEGO() As String
    getcolINTNEGO = colINTNEGO
End Function

Function getcolINTORDR() As String
    getcolINTORDR = colINTORDR
End Function

Function getcolINTSTAT() As String
    getcolINTSTAT = colINTSTAT
End Function

Function getcolINTCEXGLO() As String
    getcolINTCEXGLO = colINTCEXGLO
End Function

Function getcolINTNOOE() As String
    getcolINTNOOE = colINTNOOE
End Function

Function getcolINTFLUX() As String
    getcolINTFLUX = colINTFLUX
End Function

Function getcolINTFSTA() As String
    getcolINTFSTA = colINTFSTA
End Function

Function getcolINTLDIST() As String
    getcolINTLDIST = colINTLDIST
End Function

Function getcolINTLDNO() As String
    getcolINTLDNO = colINTLDNO
End Function

Function getcolINTETAT() As String
    getcolINTETAT = colINTETAT
End Function

Function getcolINTSITLI() As String
    getcolINTSITLI = colINTSITLI
End Function

Function getcolINTPACH() As String
    getcolINTPACH = colINTPACH
End Function

Function getcolINTCOML1() As String
    getcolINTCOML1 = colINTCOML1
End Function

Function getcolINTCLCUS() As String
    getcolINTCLCUS = colINTCLCUS
End Function

Function getcolINTURG() As String
    getcolINTURG = colINTURG
End Function

Function getcolINTEXT() As String
    getcolINTEXT = colINTEXT
End Function

Function getcolINTESCO() As String
    getcolINTESCO = colINTESCO
End Function

Function getcolINTNJESC() As String
    getcolINTNJESC = colINTNJESC
End Function

Function getcolINTPORI() As String
    getcolINTPORI = colINTPORI
End Function

Function getcolINTINCO() As String
    getcolINTINCO = colINTINCO
End Function

Function getcolINTLIEU2() As String
    getcolINTLIEU2 = colINTLIEU2
End Function

Function getcolINTTRSP() As String
    getcolINTTRSP = colINTTRSP
End Function

Function getcolINTFRAN() As String
    getcolINTFRAN = colINTFRAN
End Function

Function getcolINTVOLI() As String
    getcolINTVOLI = colINTVOLI
End Function

Function getcolINTPDSI() As String
    getcolINTPDSI = colINTPDSI
End Function

Function getcolINTTYIM() As String
    getcolINTTYIM = colINTTYIM
End Function

Function getcolINTDBAS() As String
    getcolINTDBAS = colINTDBAS
End Function

Function getcolINTDDEP() As String
    getcolINTDDEP = colINTDDEP
End Function

Function getcolINTCRED() As String
    getcolINTCRED = colINTCRED
End Function

Function getcolINTJOUR() As String
    getcolINTJOUR = colINTJOUR
End Function

Function getcolINTDARR() As String
    getcolINTDARR = colINTDARR
End Function

Function getcolINTMREG() As String
    getcolINTMREG = colINTMREG
End Function

Function getcolINTDDS() As String
    getcolINTDDS = colINTDDS
End Function

Function getcolINTNBJM() As String
    getcolINTNBJM = colINTNBJM
End Function

Function getcolINTDVAL() As String
    getcolINTDVAL = colINTDVAL
End Function

Function getcolINTDPAI() As String
    getcolINTDPAI = colINTDPAI
End Function

Function getcolINTNSEQ() As String
    getcolINTNSEQ = colINTNSEQ
End Function

Function getcolINTNLIG() As String
    getcolINTNLIG = colINTNLIG
End Function

Function getcolINTNLEN() As String
    getcolINTNLEN = colINTNLEN
End Function

Function getcolINTFICH() As String
    getcolINTFICH = colINTFICH
End Function

Function getcolINTCACT() As String
    getcolINTCACT = colINTCACT
End Function

Function getcolINTNERR() As String
    getcolINTNERR = colINTNERR
End Function

Function getcolINTMESS() As String
    getcolINTMESS = colINTMESS
End Function

Function getcolINTDCRE() As String
    getcolINTDCRE = colINTDCRE
End Function

Function getcolINTDMAJ() As String
    getcolINTDMAJ = colINTDMAJ
End Function

Function getcolINTUTIL() As String
    getcolINTUTIL = colINTUTIL
End Function

Function getcolINTDTRT() As String
    getcolINTDTRT = colINTDTRT
End Function

Function getcolINTCTVA() As String
    getcolINTCTVA = colINTCTVA
End Function

Function getcolINTUAPP() As String
    getcolINTUAPP = colINTUAPP
End Function

Function getcolINTALTF() As String
    getcolINTALTF = colINTALTF
End Function

Function getcolINTTYPUL() As String
    getcolINTTYPUL = colINTTYPUL
End Function

Function getcolINTCEXOGL() As String
    getcolINTCEXOGL = colINTCEXOGL
End Function

Function getcolINTCEXOPS() As String
    getcolINTCEXOPS = colINTCEXOPS
End Function

Function getcolINTNROUTE() As String
    getcolINTNROUTE = colINTNROUTE
End Function

Function getcolINTLIEU() As String
    getcolINTLIEU = colINTLIEU
End Function

Function getcolINTVALOF() As String
    getcolINTVALOF = colINTVALOF
End Function

Function getcolINTMOTIF() As String
    getcolINTMOTIF = colINTMOTIF
End Function

Function getcolINTTEL() As String
    getcolINTTEL = colINTTEL
End Function

Function getcolINTORI() As String
    getcolINTORI = colINTORI
End Function

Function getcolINTCSIN() As String
    getcolINTCSIN = colINTCSIN
End Function

Function getcolINTCTLA() As String
    getcolINTCTLA = colINTCTLA
End Function

Function getcolINTIRECYC() As String
    getcolINTIRECYC = colINTIRECYC
End Function

Function getcolINTCRGP() As String
    getcolINTCRGP = colINTCRGP
End Function

Function getcolINTFLIR() As String
    getcolINTFLIR = colINTFLIR
End Function

Function getcolINTNOLV() As String
    getcolINTNOLV = colINTNOLV
End Function

Function getcolINTDRAM() As String
    getcolINTDRAM = colINTDRAM
End Function

Function getcolINTPVSA() As String
    getcolINTPVSA = colINTPVSA
End Function

Function getcolINTPVSR() As String
    getcolINTPVSR = colINTPVSR
End Function

Function getcolINTPRFA() As String
    getcolINTPRFA = colINTPRFA
End Function

Function getcolINTMTDR() As String
    getcolINTMTDR = colINTMTDR
End Function

Function getcolINTMTVI() As String
    getcolINTMTVI = colINTMTVI
End Function

Function getcolINTGRA() As String
    getcolINTGRA = colINTGRA
End Function

Function getcolINTDENVREC() As String
    getcolINTDENVREC = colINTDENVREC
End Function

Function getcolINTCEAN() As String
    getcolINTCEAN = colINTCEAN
End Function

Function getcolINTCEXTJF() As String
    getcolINTCEXTJF = colINTCEXTJF
End Function

Function getcolINTEDOU() As String
    getcolINTEDOU = colINTEDOU
End Function

Function getcolINTRDOU() As String
    getcolINTRDOU = colINTRDOU
End Function

Function getcolINTDENLEV() As String
    getcolINTDENLEV = colINTDENLEV
End Function

Function getcolINTREFEXT() As String
    getcolINTREFEXT = colINTREFEXT
End Function

Function getcolINTCTRL() As String
    getcolINTCTRL = colINTCTRL
End Function

Function getcolINTFVSA() As String
    getcolINTFVSA = colINTFVSA
End Function

Function getcolINTFVSR() As String
    getcolINTFVSR = colINTFVSR
End Function

Function getcolINTCODLOG() As String
    getcolINTCODLOG = colINTCODLOG
End Function

Function getcolINTCODCAI() As String
    getcolINTCODCAI = colINTCODCAI
End Function

Function getcolINTUEREMP() As String
    getcolINTUEREMP = colINTUEREMP
End Function

Function getcolINTCINB() As String
    getcolINTCINB = colINTCINB
End Function

Function getcolINTNOLIGN() As String
    getcolINTNOLIGN = colINTNOLIGN
End Function

Function getcolINTPROPER() As String
    getcolINTPROPER = colINTPROPER
End Function

Function getrsEAN() As Integer
    getrsEAN = rsEAN
End Function

Function getrsNAZIV() As Integer
    getrsNAZIV = rsNAZIV
End Function

Function getrsINTSITE() As Integer
    getrsINTSITE = rsINTSITE
End Function

Function getrsINTQTEC() As Integer
    getrsINTQTEC = rsINTQTEC
End Function

Function getrsINTDCOM() As Integer
    getrsINTDCOM = rsINTDCOM
End Function

Function getrsINTDLIV() As Integer
    getrsINTDLIV = rsINTDLIV
End Function

Function getrsINTID() As Integer
    getrsINTID = rsINTID
End Function

Function getrsINTLCDE() As Integer
    getrsINTLCDE = rsINTLCDE
End Function

Function getrsINTCNUF() As Integer
    getrsINTCNUF = rsINTCNUF
End Function

Function getrsINTCCOM() As Integer
    getrsINTCCOM = rsINTCCOM
End Function

Function getrsINTNFILF() As Integer
    getrsINTNFILF = rsINTNFILF
End Function

Function getrsINTFILC() As Integer
    getrsINTFILC = rsINTFILC
End Function

Function getrsINTCONF() As Integer
    getrsINTCONF = rsINTCONF
End Function

Function getrsINTGREL() As Integer
    getrsINTGREL = rsINTGREL
End Function

Function getrsINTDEVI() As Integer
    getrsINTDEVI = rsINTDEVI
End Function

Function getrsINTCOUC() As Integer
    getrsINTCOUC = rsINTCOUC
End Function

Function getrsINTTXCH() As Integer
    getrsINTTXCH = rsINTTXCH
End Function

Function getrsINTCOM1() As Integer
    getrsINTCOM1 = rsINTCOM1
End Function

Function getrsINTCOM2() As Integer
    getrsINTCOM2 = rsINTCOM2
End Function

Function getrsINTENLEV() As Integer
    getrsINTENLEV = rsINTENLEV
End Function

Function getrsINTDLIM() As Integer
    getrsINTDLIM = rsINTDLIM
End Function

Function getrsINTCODE() As Integer
    getrsINTCODE = rsINTCODE
End Function

Function getrsINTRCOM() As Integer
    getrsINTRCOM = rsINTRCOM
End Function

Function getrsINTCEXVA() As Integer
    getrsINTCEXVA = rsINTCEXVA
End Function

Function getrsINTCEXVL() As Integer
    getrsINTCEXVL = rsINTCEXVL
End Function

Function getrsINTUAUVC() As Integer
    getrsINTUAUVC = rsINTUAUVC
End Function

Function getrsINTNEGO() As Integer
    getrsINTNEGO = rsINTNEGO
End Function

Function getrsINTORDR() As Integer
    getrsINTORDR = rsINTORDR
End Function

Function getrsINTSTAT() As Integer
    getrsINTSTAT = rsINTSTAT
End Function

Function getrsINTCEXGLO() As Integer
    getrsINTCEXGLO = rsINTCEXGLO
End Function

Function getrsINTNOOE() As Integer
    getrsINTNOOE = rsINTNOOE
End Function

Function getrsINTFLUX() As Integer
    getrsINTFLUX = rsINTFLUX
End Function

Function getrsINTFSTA() As Integer
    getrsINTFSTA = rsINTFSTA
End Function

Function getrsINTLDIST() As Integer
    getrsINTLDIST = rsINTLDIST
End Function

Function getrsINTLDNO() As Integer
    getrsINTLDNO = rsINTLDNO
End Function

Function getrsINTETAT() As Integer
    getrsINTETAT = rsINTETAT
End Function

Function getrsINTSITLI() As Integer
    getrsINTSITLI = rsINTSITLI
End Function

Function getrsINTPACH() As Integer
    getrsINTPACH = rsINTPACH
End Function

Function getrsINTCOML1() As Integer
    getrsINTCOML1 = rsINTCOML1
End Function

Function getrsINTCLCUS() As Integer
    getrsINTCLCUS = rsINTCLCUS
End Function

Function getrsINTURG() As Integer
    getrsINTURG = rsINTURG
End Function

Function getrsINTEXT() As Integer
    getrsINTEXT = rsINTEXT
End Function

Function getrsINTESCO() As Integer
    getrsINTESCO = rsINTESCO
End Function

Function getrsINTNJESC() As Integer
    getrsINTNJESC = rsINTNJESC
End Function

Function getrsINTPORI() As Integer
    getrsINTPORI = rsINTPORI
End Function

Function getrsINTINCO() As Integer
    getrsINTINCO = rsINTINCO
End Function

Function getrsINTLIEU2() As Integer
    getrsINTLIEU2 = rsINTLIEU2
End Function

Function getrsINTTRSP() As Integer
    getrsINTTRSP = rsINTTRSP
End Function

Function getrsINTFRAN() As Integer
    getrsINTFRAN = rsINTFRAN
End Function

Function getrsINTVOLI() As Integer
    getrsINTVOLI = rsINTVOLI
End Function

Function getrsINTPDSI() As Integer
    getrsINTPDSI = rsINTPDSI
End Function

Function getrsINTTYIM() As Integer
    getrsINTTYIM = rsINTTYIM
End Function

Function getrsINTDBAS() As Integer
    getrsINTDBAS = rsINTDBAS
End Function

Function getrsINTDDEP() As Integer
    getrsINTDDEP = rsINTDDEP
End Function

Function getrsINTCRED() As Integer
    getrsINTCRED = rsINTCRED
End Function

Function getrsINTJOUR() As Integer
    getrsINTJOUR = rsINTJOUR
End Function

Function getrsINTDARR() As Integer
    getrsINTDARR = rsINTDARR
End Function

Function getrsINTMREG() As Integer
    getrsINTMREG = rsINTMREG
End Function

Function getrsINTDDS() As Integer
    getrsINTDDS = rsINTDDS
End Function

Function getrsINTNBJM() As Integer
    getrsINTNBJM = rsINTNBJM
End Function

Function getrsINTDVAL() As Integer
    getrsINTDVAL = rsINTDVAL
End Function

Function getrsINTDPAI() As Integer
    getrsINTDPAI = rsINTDPAI
End Function

Function getrsINTNSEQ() As Integer
    getrsINTNSEQ = rsINTNSEQ
End Function

Function getrsINTNLIG() As Integer
    getrsINTNLIG = rsINTNLIG
End Function

Function getrsINTNLEN() As Integer
    getrsINTNLEN = rsINTNLEN
End Function

Function getrsINTFICH() As Integer
    getrsINTFICH = rsINTFICH
End Function

Function getrsINTCACT() As Integer
    getrsINTCACT = rsINTCACT
End Function

Function getrsINTNERR() As Integer
    getrsINTNERR = rsINTNERR
End Function

Function getrsINTMESS() As Integer
    getrsINTMESS = rsINTMESS
End Function

Function getrsINTDCRE() As Integer
    getrsINTDCRE = rsINTDCRE
End Function

Function getrsINTDMAJ() As Integer
    getrsINTDMAJ = rsINTDMAJ
End Function

Function getrsINTUTIL() As Integer
    getrsINTUTIL = rsINTUTIL
End Function

Function getrsINTDTRT() As Integer
    getrsINTDTRT = rsINTDTRT
End Function

Function getrsINTCTVA() As Integer
    getrsINTCTVA = rsINTCTVA
End Function

Function getrsINTUAPP() As Integer
    getrsINTUAPP = rsINTUAPP
End Function

Function getrsINTALTF() As Integer
    getrsINTALTF = rsINTALTF
End Function

Function getrsINTTYPUL() As Integer
    getrsINTTYPUL = rsINTTYPUL
End Function

Function getrsINTCEXOGL() As Integer
    getrsINTCEXOGL = rsINTCEXOGL
End Function

Function getrsINTCEXOPS() As Integer
    getrsINTCEXOPS = rsINTCEXOPS
End Function

Function getrsINTNROUTE() As Integer
    getrsINTNROUTE = rsINTNROUTE
End Function

Function getrsINTLIEU() As Integer
    getrsINTLIEU = rsINTLIEU
End Function

Function getrsINTVALOF() As Integer
    getrsINTVALOF = rsINTVALOF
End Function

Function getrsINTMOTIF() As Integer
    getrsINTMOTIF = rsINTMOTIF
End Function

Function getrsINTTEL() As Integer
    getrsINTTEL = rsINTTEL
End Function

Function getrsINTORI() As Integer
    getrsINTORI = rsINTORI
End Function

Function getrsINTCSIN() As Integer
    getrsINTCSIN = rsINTCSIN
End Function

Function getrsINTCTLA() As Integer
    getrsINTCTLA = rsINTCTLA
End Function

Function getrsINTIRECYC() As Integer
    getrsINTIRECYC = rsINTIRECYC
End Function

Function getrsINTCRGP() As Integer
    getrsINTCRGP = rsINTCRGP
End Function

Function getrsINTFLIR() As Integer
    getrsINTFLIR = rsINTFLIR
End Function

Function getrsINTNOLV() As Integer
    getrsINTNOLV = rsINTNOLV
End Function

Function getrsINTDRAM() As Integer
    getrsINTDRAM = rsINTDRAM
End Function

Function getrsINTPVSA() As Integer
    getrsINTPVSA = rsINTPVSA
End Function

Function getrsINTPVSR() As Integer
    getrsINTPVSR = rsINTPVSR
End Function

Function getrsINTPRFA() As Integer
    getrsINTPRFA = rsINTPRFA
End Function

Function getrsINTMTDR() As Integer
    getrsINTMTDR = rsINTMTDR
End Function

Function getrsINTMTVI() As Integer
    getrsINTMTVI = rsINTMTVI
End Function

Function getrsINTGRA() As Integer
    getrsINTGRA = rsINTGRA
End Function

Function getrsINTDENVREC() As Integer
    getrsINTDENVREC = rsINTDENVREC
End Function

Function getrsINTCEAN() As Integer
    getrsINTCEAN = rsINTCEAN
End Function

Function getrsINTCEXTJF() As Integer
    getrsINTCEXTJF = rsINTCEXTJF
End Function

Function getrsINTEDOU() As Integer
    getrsINTEDOU = rsINTEDOU
End Function

Function getrsINTRDOU() As Integer
    getrsINTRDOU = rsINTRDOU
End Function

Function getrsINTDENLEV() As Integer
    getrsINTDENLEV = rsINTDENLEV
End Function

Function getrsINTREFEXT() As Integer
    getrsINTREFEXT = rsINTREFEXT
End Function

Function getrsINTCTRL() As Integer
    getrsINTCTRL = rsINTCTRL
End Function

Function getrsINTFVSA() As Integer
    getrsINTFVSA = rsINTFVSA
End Function

Function getrsINTFVSR() As Integer
    getrsINTFVSR = rsINTFVSR
End Function

Function getrsINTCODLOG() As Integer
    getrsINTCODLOG = rsINTCODLOG
End Function

Function getrsINTCODCAI() As Integer
    getrsINTCODCAI = rsINTCODCAI
End Function

Function getrsINTUEREMP() As Integer
    getrsINTUEREMP = rsINTUEREMP
End Function

Function getrsINTCINB() As Integer
    getrsINTCINB = rsINTCINB
End Function

Function getrsINTNOLIGN() As Integer
    getrsINTNOLIGN = rsINTNOLIGN
End Function

Function getrsINTPROPER() As Integer
    getrsINTPROPER = rsINTPROPER
End Function


