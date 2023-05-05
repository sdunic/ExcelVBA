Attribute VB_Name = "cfg"
Option Explicit
'config dokument u kojemu se:
' - inicijalno generira db connection string
' - parametriziraju stupci u tablici
' - parametriziraju indeksi recordseta kojeg uèitavamo iz Oracle baze podataka

Dim cTNUMSGID, cTNULNLIG, cTNUCNUF, cTNUSUPDESC, cTNUCCOM, cTNUAGRP, cTNUCEXR, cTNUADESC, cTNULV, cTNULU, cTNUSITE, cTNUSDESC, cTNUPACH, cTNUUAPP, cTNUNNC, cTNUEXNNC, cTNUPADDEB, cTNUPADFIN, cTNUTCP As String
Dim cTNUVAL601, cTNUUAPP601, cTNUDDEB601, cTNUDFIN601, cTNUVAL602, cTNUUAPP602, cTNUDDEB602, cTNUDFIN602, cTNUVAL603, cTNUUAPP603, cTNUDDEB603, cTNUDFIN603, cTNUVAL604, cTNUUAPP604, cTNUDDEB604, cTNUDFIN604, cTNUVAL605, cTNUUAPP605, cTNUDDEB605, cTNUDFIN605, cTNUVAL606, cTNUUAPP606, cTNUDDEB606, cTNUDFIN606 As String
Dim cARCCODE, cPRINCIPAL, cASORTIMAN As String
Dim cBROJPROMJENA As String

Dim rTNUMSGID, rTNULNLIG, rTNUCNUF, rTNUSUPDESC, rTNUCCOM, rTNUAGRP, rTNUCEXR, rTNUADESC, rTNULV, rTNULU, rTNUSITE, rTNUSDESC, rTNUPACH, rTNUPASTPACH, rTNUFUTPACH, rTNUUAPP, rTNUNNC, rTNUEXNNC, rTNUPADDEB, rTNUPADFIN, rTNUTCP As Integer
Dim rTNUVAL601, rTNUUAPP601, rTNUDDEB601, rTNUDFIN601, rTNUVAL602, rTNUUAPP602, rTNUDDEB602, rTNUDFIN602, rTNUVAL603, rTNUUAPP603, rTNUDDEB603, rTNUDFIN603, rTNUVAL604, rTNUUAPP604, rTNUDDEB604, rTNUDFIN604, rTNUVAL605, rTNUUAPP605, rTNUDDEB605, rTNUDFIN605, rTNUVAL606, rTNUUAPP606, rTNUDDEB606, rTNUDFIN606 As Integer
Dim rARCCODE, rPRINCIPAL, rASORTIMAN As Integer

Dim rTNUFUT601, rTNUFUT602, rTNUFUT603, rTNUFUT604, rTNUFUT605, rTNUFUT606, rTNUPAST601, rTNUPAST602, rTNUPAST603, rTNUPAST604, rTNUPAST605, rTNUPAST606 As Integer

Public oldValues As Collection
Public keys As Collection

Sub InitCollections()
    Set oldValues = New Collection
    Set keys = New Collection
End Sub

Sub Init()
    'parametrizacija stupaca
    
    cTNUMSGID = "B"
    cTNULNLIG = "C"
    cTNUCNUF = "D"
    cTNUSUPDESC = "E"
    cTNUCCOM = "F"
    cTNUAGRP = "G"
    cTNUCEXR = "H"
    cARCCODE = "I"
    cTNUADESC = "J"
    cPRINCIPAL = "K"
    cASORTIMAN = "L"
    cTNULV = "M"
    cTNULU = "N"
    cTNUSITE = "O"
    cTNUSDESC = "P"
    cTNUPACH = "Q"
    cTNUUAPP = "R"
    cTNUNNC = "S"
    cTNUEXNNC = "T"
    cTNUPADDEB = "U"
    cTNUPADFIN = "V"
    cTNUTCP = "W"
    cTNUVAL601 = "X"
    cTNUUAPP601 = "Y"
    cTNUDDEB601 = "Z"
    cTNUDFIN601 = "AA"
    cTNUVAL602 = "AB"
    cTNUUAPP602 = "AC"
    cTNUDDEB602 = "AD"
    cTNUDFIN602 = "AE"
    cTNUVAL603 = "AF"
    cTNUUAPP603 = "AG"
    cTNUDDEB603 = "AH"
    cTNUDFIN603 = "AI"
    cTNUVAL604 = "AJ"
    cTNUUAPP604 = "AK"
    cTNUDDEB604 = "AL"
    cTNUDFIN604 = "AM"
    cTNUVAL605 = "AN"
    cTNUUAPP605 = "AO"
    cTNUDDEB605 = "AP"
    cTNUDFIN605 = "AQ"
    cTNUVAL606 = "AR"
    cTNUUAPP606 = "AS"
    cTNUDDEB606 = "AT"
    cTNUDFIN606 = "AU"
    cBROJPROMJENA = "AV"
    
    rTNUMSGID = 0
    rTNULNLIG = 1
    rTNUCNUF = 2
    rTNUSUPDESC = 3
    rTNUCCOM = 4
    rTNUAGRP = 5
    rTNUCEXR = 6
    rTNUADESC = 7
    rTNULV = 8
    rTNULU = 9
    rTNUSITE = 10
    rTNUSDESC = 11
    rTNUPACH = 12
    rTNUPASTPACH = 13
    rTNUFUTPACH = 14
    rTNUUAPP = 15
    rTNUNNC = 16
    rTNUEXNNC = 17
    rTNUPADDEB = 18
    rTNUPADFIN = 19
    rTNUTCP = 20
    rTNUVAL601 = 21
    rTNUUAPP601 = 22
    rTNUDDEB601 = 23
    rTNUDFIN601 = 24
    rTNUPAST601 = 25
    rTNUFUT601 = 26
    rTNUVAL602 = 27
    rTNUUAPP602 = 28
    rTNUDDEB602 = 29
    rTNUDFIN602 = 30
    rTNUPAST602 = 31
    rTNUFUT602 = 32
    rTNUVAL603 = 33
    rTNUUAPP603 = 34
    rTNUDDEB603 = 35
    rTNUDFIN603 = 36
    rTNUPAST603 = 37
    rTNUFUT603 = 38
    rTNUVAL604 = 39
    rTNUUAPP604 = 40
    rTNUDDEB604 = 41
    rTNUDFIN604 = 42
    rTNUPAST604 = 43
    rTNUFUT604 = 44
    rTNUVAL605 = 45
    rTNUUAPP605 = 46
    rTNUDDEB605 = 47
    rTNUDFIN605 = 48
    rTNUPAST605 = 49
    rTNUFUT605 = 50
    rTNUVAL606 = 51
    rTNUUAPP606 = 52
    rTNUDDEB606 = 53
    rTNUDFIN606 = 54
    rTNUPAST606 = 55
    rTNUFUT606 = 56

    rARCCODE = 70
    rPRINCIPAL = 71
    rASORTIMAN = 72
    
End Sub

Sub addKeyValue(key As String, val As String)
    oldValues.Add val, key
End Sub

Sub addKeyItem(key As String)
    keys.Add key
End Sub

Function getValueByKey(key As String) As String
    getValueByKey = oldValues(key)
End Function

Function getcBROJPROMJENA() As String
    getcBROJPROMJENA = cBROJPROMJENA
End Function

Function getcTNUMSGID() As String
    getcTNUMSGID = cTNUMSGID
End Function

Function getcTNULNLIG() As String
    getcTNULNLIG = cTNULNLIG
End Function

Function getcTNUCNUF() As String
    getcTNUCNUF = cTNUCNUF
End Function

Function getcTNUSUPDESC() As String
    getcTNUSUPDESC = cTNUSUPDESC
End Function

Function getcTNUCCOM() As String
    getcTNUCCOM = cTNUCCOM
End Function

Function getcTNUAGRP() As String
    getcTNUAGRP = cTNUAGRP
End Function

Function getcTNUCEXR() As String
    getcTNUCEXR = cTNUCEXR
End Function

Function getcARCCODE() As String
    getcARCCODE = cARCCODE
End Function

Function getcTNUADESC() As String
    getcTNUADESC = cTNUADESC
End Function

Function getcPRINCIPAL() As String
    getcPRINCIPAL = cPRINCIPAL
End Function

Function getcASORTIMAN() As String
    getcASORTIMAN = cASORTIMAN
End Function

Function getcTNULV() As String
    getcTNULV = cTNULV
End Function

Function getcTNULU() As String
    getcTNULU = cTNULU
End Function

Function getcTNUSITE() As String
    getcTNUSITE = cTNUSITE
End Function

Function getcTNUSDESC() As String
    getcTNUSDESC = cTNUSDESC
End Function

Function getcTNUPACH() As String
    getcTNUPACH = cTNUPACH
End Function

Function getcTNUUAPP() As String
    getcTNUUAPP = cTNUUAPP
End Function

Function getcTNUNNC() As String
    getcTNUNNC = cTNUNNC
End Function

Function getcTNUEXNNC() As String
    getcTNUEXNNC = cTNUEXNNC
End Function

Function getcTNUPADDEB() As String
    getcTNUPADDEB = cTNUPADDEB
End Function

Function getcTNUPADFIN() As String
    getcTNUPADFIN = cTNUPADFIN
End Function

Function getcTNUTCP() As String
    getcTNUTCP = cTNUTCP
End Function

Function getcTNUVAL601() As String
    getcTNUVAL601 = cTNUVAL601
End Function

Function getcTNUUAPP601() As String
    getcTNUUAPP601 = cTNUUAPP601
End Function

Function getcTNUDDEB601() As String
    getcTNUDDEB601 = cTNUDDEB601
End Function

Function getcTNUDFIN601() As String
    getcTNUDFIN601 = cTNUDFIN601
End Function

Function getcTNUVAL602() As String
    getcTNUVAL602 = cTNUVAL602
End Function

Function getcTNUUAPP602() As String
    getcTNUUAPP602 = cTNUUAPP602
End Function

Function getcTNUDDEB602() As String
    getcTNUDDEB602 = cTNUDDEB602
End Function

Function getcTNUDFIN602() As String
    getcTNUDFIN602 = cTNUDFIN602
End Function

Function getcTNUVAL603() As String
    getcTNUVAL603 = cTNUVAL603
End Function

Function getcTNUUAPP603() As String
    getcTNUUAPP603 = cTNUUAPP603
End Function

Function getcTNUDDEB603() As String
    getcTNUDDEB603 = cTNUDDEB603
End Function

Function getcTNUDFIN603() As String
    getcTNUDFIN603 = cTNUDFIN603
End Function

Function getcTNUVAL604() As String
    getcTNUVAL604 = cTNUVAL604
End Function

Function getcTNUUAPP604() As String
    getcTNUUAPP604 = cTNUUAPP604
End Function

Function getcTNUDDEB604() As String
    getcTNUDDEB604 = cTNUDDEB604
End Function

Function getcTNUDFIN604() As String
    getcTNUDFIN604 = cTNUDFIN604
End Function

Function getcTNUVAL605() As String
    getcTNUVAL605 = cTNUVAL605
End Function

Function getcTNUUAPP605() As String
    getcTNUUAPP605 = cTNUUAPP605
End Function

Function getcTNUDDEB605() As String
    getcTNUDDEB605 = cTNUDDEB605
End Function

Function getcTNUDFIN605() As String
    getcTNUDFIN605 = cTNUDFIN605
End Function

Function getcTNUVAL606() As String
    getcTNUVAL606 = cTNUVAL606
End Function

Function getcTNUUAPP606() As String
    getcTNUUAPP606 = cTNUUAPP606
End Function

Function getcTNUDDEB606() As String
    getcTNUDDEB606 = cTNUDDEB606
End Function

Function getcTNUDFIN606() As String
    getcTNUDFIN606 = cTNUDFIN606
End Function


Function getrTNUMSGID() As Integer
    getrTNUMSGID = rTNUMSGID
End Function

Function getrTNULNLIG() As Integer
    getrTNULNLIG = rTNULNLIG
End Function

Function getrTNUCNUF() As Integer
    getrTNUCNUF = rTNUCNUF
End Function

Function getrTNUSUPDESC() As Integer
    getrTNUSUPDESC = rTNUSUPDESC
End Function

Function getrTNUCCOM() As Integer
    getrTNUCCOM = rTNUCCOM
End Function

Function getrTNUAGRP() As Integer
    getrTNUAGRP = rTNUAGRP
End Function

Function getrTNUCEXR() As Integer
    getrTNUCEXR = rTNUCEXR
End Function

Function getrTNUADESC() As Integer
    getrTNUADESC = rTNUADESC
End Function

Function getrTNULV() As Integer
    getrTNULV = rTNULV
End Function

Function getrTNULU() As Integer
    getrTNULU = rTNULU
End Function

Function getrTNUSITE() As Integer
    getrTNUSITE = rTNUSITE
End Function

Function getrTNUSDESC() As Integer
    getrTNUSDESC = rTNUSDESC
End Function

Function getrTNUPACH() As Integer
    getrTNUPACH = rTNUPACH
End Function

Function getrTNUPASTPACH() As Integer
    getrTNUPASTPACH = rTNUPASTPACH
End Function

Function getrTNUFUTPACH() As Integer
    getrTNUFUTPACH = rTNUFUTPACH
End Function

Function getrTNUUAPP() As Integer
    getrTNUUAPP = rTNUUAPP
End Function

Function getrTNUNNC() As Integer
    getrTNUNNC = rTNUNNC
End Function

Function getrTNUEXNNC() As Integer
    getrTNUEXNNC = rTNUEXNNC
End Function

Function getrTNUPADDEB() As Integer
    getrTNUPADDEB = rTNUPADDEB
End Function

Function getrTNUPADFIN() As Integer
    getrTNUPADFIN = rTNUPADFIN
End Function

Function getrTNUTCP() As Integer
    getrTNUTCP = rTNUTCP
End Function

Function getrTNUVAL601() As Integer
    getrTNUVAL601 = rTNUVAL601
End Function

Function getrTNUUAPP601() As Integer
    getrTNUUAPP601 = rTNUUAPP601
End Function

Function getrTNUDDEB601() As Integer
    getrTNUDDEB601 = rTNUDDEB601
End Function

Function getrTNUDFIN601() As Integer
    getrTNUDFIN601 = rTNUDFIN601
End Function

Function getrTNUVAL602() As Integer
    getrTNUVAL602 = rTNUVAL602
End Function

Function getrTNUUAPP602() As Integer
    getrTNUUAPP602 = rTNUUAPP602
End Function

Function getrTNUDDEB602() As Integer
    getrTNUDDEB602 = rTNUDDEB602
End Function

Function getrTNUDFIN602() As Integer
    getrTNUDFIN602 = rTNUDFIN602
End Function

Function getrTNUVAL603() As Integer
    getrTNUVAL603 = rTNUVAL603
End Function

Function getrTNUUAPP603() As Integer
    getrTNUUAPP603 = rTNUUAPP603
End Function

Function getrTNUDDEB603() As Integer
    getrTNUDDEB603 = rTNUDDEB603
End Function

Function getrTNUDFIN603() As Integer
    getrTNUDFIN603 = rTNUDFIN603
End Function

Function getrTNUVAL604() As Integer
    getrTNUVAL604 = rTNUVAL604
End Function

Function getrTNUUAPP604() As Integer
    getrTNUUAPP604 = rTNUUAPP604
End Function

Function getrTNUDDEB604() As Integer
    getrTNUDDEB604 = rTNUDDEB604
End Function

Function getrTNUDFIN604() As Integer
    getrTNUDFIN604 = rTNUDFIN604
End Function

Function getrTNUVAL605() As Integer
    getrTNUVAL605 = rTNUVAL605
End Function

Function getrTNUUAPP605() As Integer
    getrTNUUAPP605 = rTNUUAPP605
End Function

Function getrTNUDDEB605() As Integer
    getrTNUDDEB605 = rTNUDDEB605
End Function

Function getrTNUDFIN605() As Integer
    getrTNUDFIN605 = rTNUDFIN605
End Function

Function getrTNUVAL606() As Integer
    getrTNUVAL606 = rTNUVAL606
End Function

Function getrTNUUAPP606() As Integer
    getrTNUUAPP606 = rTNUUAPP606
End Function

Function getrTNUDDEB606() As Integer
    getrTNUDDEB606 = rTNUDDEB606
End Function

Function getrTNUDFIN606() As Integer
    getrTNUDFIN606 = rTNUDFIN606
End Function

Function getrARCCODE() As Integer
    getrARCCODE = rARCCODE
End Function

Function getrPRINCIPAL() As Integer
    getrPRINCIPAL = rPRINCIPAL
End Function

Function getrASORTIMAN() As Integer
    getrASORTIMAN = rASORTIMAN
End Function


Function getrTNUFUT601() As Integer
    getrTNUFUT601 = rTNUFUT601
End Function
Function getrTNUFUT602() As Integer
    getrTNUFUT602 = rTNUFUT602
End Function
Function getrTNUFUT603() As Integer
    getrTNUFUT603 = rTNUFUT603
End Function
Function getrTNUFUT604() As Integer
    getrTNUFUT604 = rTNUFUT604
End Function
Function getrTNUFUT605() As Integer
    getrTNUFUT605 = rTNUFUT605
End Function
Function getrTNUFUT606() As Integer
    getrTNUFUT606 = rTNUFUT606
End Function
Function getrTNUPAST601() As Integer
    getrTNUPAST601 = rTNUPAST601
End Function
Function getrTNUPAST602() As Integer
    getrTNUPAST602 = rTNUPAST602
End Function
Function getrTNUPAST603() As Integer
    getrTNUPAST603 = rTNUPAST603
End Function
Function getrTNUPAST604() As Integer
    getrTNUPAST604 = rTNUPAST604
End Function
Function getrTNUPAST605() As Integer
    getrTNUPAST605 = rTNUPAST605
End Function
Function getrTNUPAST606() As Integer
    getrTNUPAST606 = rTNUPAST606
End Function

