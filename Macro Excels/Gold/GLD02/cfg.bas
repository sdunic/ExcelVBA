Attribute VB_Name = "cfg"
'config dokument u kojemu se:
' - inicijalno generira db connection string
' - parametriziraju stupci u tablici
' - parametriziraju indeksi recordseta kojeg uèitavamo iz Oracle baze podataka
Dim colSifraArtikla, colBarkodArtikla, colNazivArtikla, colBrand, colPrincipal As String
Dim colNivo1, colNaziv1, colNivo2, colNaziv2, colNivo3, colNaziv3, colNivo4, colNaziv4, colNivo5, colNaziv5 As String
Dim colAsortiman, colTSC, colOpis, colSvojstva, colPocetnaCijena As String
Dim colMPC_ADatum, colMPC_ACijena, colMPC_ANovaCijena, colMPC_AIndeks, colMPC_BDatum, colMPC_BCijena, colMPC_BNovaCijena, colMPC_BIndeks, colMPC_CDatum, colMPC_CCijena, colMPC_CNovaCijena, colMPC_CIndeks, colMPC_DDatum, colMPC_DCijena, colMPC_DNovaCijena, colMPC_DIndeks As String
Dim colMPC_S1Datum, colMPC_S1Cijena, colMPC_S1NovaCijena, colMPC_S1Indeks As String
Dim colMPC_S2Datum, colMPC_S2Cijena, colMPC_S2NovaCijena, colMPC_S2Indeks As String
Dim colMPC_S3Datum, colMPC_S3Cijena, colMPC_S3NovaCijena, colMPC_S3Indeks As String
Dim colMPC_KAMPDatum, colMPC_KAMPCijena, colMPC_KAMPNovaCijena, colMPC_KAMPIndeks As String
Dim colRedak, colPoreznaGrupa, colCEXV, colBrojPromjena As String

Dim rsSifraArtikla, rsBarkodArtikla, rsNazivArtikla, rsBrand, rsPrincipal As Integer
Dim rsNivo1, rsNaziv1, rsNivo2, rsNaziv2, rsNivo3, rsNaziv3, rsNivo4, rsNaziv4, rsNivo5, rsNaziv5 As Integer
Dim rsAsortiman, rsTSC, rsOpis, rsSvojstva As Integer
Dim rsMPC_ANtar, rsMPC_ACijena, rsMPC_ADatum, rsMPC_BNtar, rsMPC_BCijena, rsMPC_BDatum, rsMPC_CNtar, rsMPC_CCijena, rsMPC_CDatum, rsMPC_DNtar, rsMPC_DCijena, rsMPC_DDatum As Integer
Dim rsMPC_S1Ntar, rsMPC_S1Cijena, rsMPC_S1Datum As Integer
Dim rsMPC_S2Ntar, rsMPC_S2Cijena, rsMPC_S2Datum As Integer
Dim rsMPC_S3Ntar, rsMPC_S3Cijena, rsMPC_S3Datum As Integer
Dim rsMPC_KAMPNtar, rsMPC_KAMPCijena, rsMPC_KAMPDatum As Integer
Dim rsPoreznaGrupa, rsCEXV As Integer


Sub Init()

    colSifraArtikla = "B"
    colBarkodArtikla = "C"
    colNazivArtikla = "D"
    colBrand = "E"
    colPrincipal = "F"
    colNivo1 = "G"
    colNaziv1 = "H"
    colNivo2 = "I"
    colNaziv2 = "J"
    colNivo3 = "K"
    colNaziv3 = "L"
    colNivo4 = "M"
    colNaziv4 = "N"
    colNivo5 = "O"
    colNaziv5 = "P"
    colAsortiman = "Q"
    colTSC = "R"
    colOpis = "S"
    colSvojstva = "T"
    colPocetnaCijena = "U"
    colMPC_ADatum = "V"
    colMPC_ACijena = "W"
    colMPC_ANovaCijena = "X"
    colMPC_AIndeks = "Y"
    colMPC_BDatum = "Z"
    colMPC_BCijena = "AA"
    colMPC_BNovaCijena = "AB"
    colMPC_BIndeks = "AC"
    colMPC_CDatum = "AD"
    colMPC_CCijena = "AE"
    colMPC_CNovaCijena = "AF"
    colMPC_CIndeks = "AG"
    colMPC_DDatum = "AH"
    colMPC_DCijena = "AI"
    colMPC_DNovaCijena = "AJ"
    colMPC_DIndeks = "AK"
    colMPC_S1Datum = "AL"
    colMPC_S1Cijena = "AM"
    colMPC_S1NovaCijena = "AN"
    colMPC_S1Indeks = "AO"
    
    
    colMPC_S2Datum = "AP"
    colMPC_S2Cijena = "AQ"
    colMPC_S2NovaCijena = "AR"
    colMPC_S2Indeks = "AS"
    
    colMPC_S3Datum = "AT"
    colMPC_S3Cijena = "AU"
    colMPC_S3NovaCijena = "AV"
    colMPC_S3Indeks = "AW"
    
    colMPC_KAMPDatum = "AX"
    colMPC_KAMPCijena = "AY"
    colMPC_KAMPNovaCijena = "AZ"
    colMPC_KAMPIndeks = "BA"
    
    colRedak = "BB"
    colPoreznaGrupa = "BC"
    colCEXV = "BD"
    colBrojPromjena = "BE"

    rsSifraArtikla = 0
    rsBarkodArtikla = 1
    rsNazivArtikla = 2
    rsBrand = 3
    rsPrincipal = 4
    rsNivo1 = 5
    rsNaziv1 = 6
    rsNivo2 = 7
    rsNaziv2 = 8
    rsNivo3 = 9
    rsNaziv3 = 10
    rsNivo4 = 11
    rsNaziv4 = 12
    rsNivo5 = 13
    rsNaziv5 = 14
    rsAsortiman = 15
    rsTSC = 16
    rsOpis = 17
    rsSvojstva = 18
    rsMPC_ANtar = 19
    rsMPC_ACijena = 20
    rsMPC_ADatum = 21
    rsMPC_BNtar = 22
    rsMPC_BCijena = 23
    rsMPC_BDatum = 24
    rsMPC_CNtar = 25
    rsMPC_CCijena = 26
    rsMPC_CDatum = 27
    rsMPC_DNtar = 28
    rsMPC_DCijena = 29
    rsMPC_DDatum = 30
    rsMPC_S1Ntar = 31
    rsMPC_S1Cijena = 32
    rsMPC_S1Datum = 33
    rsMPC_S2Ntar = 34
    rsMPC_S2Cijena = 35
    rsMPC_S2Datum = 36
    rsMPC_S3Ntar = 37
    rsMPC_S3Cijena = 38
    rsMPC_S3Datum = 39
    rsMPC_KAMPNtar = 40
    rsMPC_KAMPCijena = 41
    rsMPC_KAMPDatum = 42
    rsPoreznaGrupa = 43
    rsCEXV = 44

End Sub

Function getRsSifraArtikla() As Integer
    getRsSifraArtikla = rsSifraArtikla
End Function

Function getRsBarkodArtikla() As Integer
    getRsBarkodArtikla = rsBarkodArtikla
End Function

Function getRsNazivArtikla() As Integer
    getRsNazivArtikla = rsNazivArtikla
End Function

Function getRsBrand() As Integer
    getRsBrand = rsBrand
End Function

Function getRsPrincipal() As Integer
    getRsPrincipal = rsPrincipal
End Function

Function getRsNivo1() As Integer
    getRsNivo1 = rsNivo1
End Function

Function getRsNaziv1() As Integer
    getRsNaziv1 = rsNaziv1
End Function

Function getRsNivo2() As Integer
    getRsNivo2 = rsNivo2
End Function

Function getRsNaziv2() As Integer
    getRsNaziv2 = rsNaziv2
End Function

Function getRsNivo3() As Integer
    getRsNivo3 = rsNivo3
End Function

Function getRsNaziv3() As Integer
    getRsNaziv3 = rsNaziv3
End Function

Function getRsNivo4() As Integer
    getRsNivo4 = rsNivo4
End Function

Function getRsNaziv4() As Integer
    getRsNaziv4 = rsNaziv4
End Function

Function getRsNivo5() As Integer
    getRsNivo5 = rsNivo5
End Function

Function getRsNaziv5() As Integer
    getRsNaziv5 = rsNaziv5
End Function

Function getRsAsortiman() As Integer
    getRsAsortiman = rsAsortiman
End Function

Function getRsTSC() As Integer
    getRsTSC = rsTSC
End Function

Function getRsOpis() As Integer
    getRsOpis = rsOpis
End Function

Function getRsSvojstva() As Integer
    getRsSvojstva = rsSvojstva
End Function

Function getRsMPC_ANtar() As Integer
    getRsMPC_ANtar = rsMPC_ANtar
End Function

Function getRsMPC_ACijena() As Integer
    getRsMPC_ACijena = rsMPC_ACijena
End Function

Function getRsMPC_ADatum() As Integer
    getRsMPC_ADatum = rsMPC_ADatum
End Function

Function getRsMPC_BNtar() As Integer
    getRsMPC_BNtar = rsMPC_BNtar
End Function

Function getRsMPC_BCijena() As Integer
    getRsMPC_BCijena = rsMPC_BCijena
End Function

Function getRsMPC_BDatum() As Integer
    getRsMPC_BDatum = rsMPC_BDatum
End Function

Function getRsMPC_CNtar() As Integer
    getRsMPC_CNtar = rsMPC_CNtar
End Function

Function getRsMPC_CCijena() As Integer
    getRsMPC_CCijena = rsMPC_CCijena
End Function

Function getRsMPC_CDatum() As Integer
    getRsMPC_CDatum = rsMPC_CDatum
End Function

Function getRsMPC_DNtar() As Integer
    getRsMPC_DNtar = rsMPC_DNtar
End Function

Function getRsMPC_DCijena() As Integer
    getRsMPC_DCijena = rsMPC_DCijena
End Function

Function getRsMPC_DDatum() As Integer
    getRsMPC_DDatum = rsMPC_DDatum
End Function

Function getRsMPC_S1Ntar() As Integer
    getRsMPC_S1Ntar = rsMPC_S1Ntar
End Function

Function getRsMPC_S1Cijena() As Integer
    getRsMPC_S1Cijena = rsMPC_S1Cijena
End Function

Function getRsMPC_S1Datum() As Integer
    getRsMPC_S1Datum = rsMPC_S1Datum
End Function

Function getRsMPC_S2Ntar() As Integer
    getRsMPC_S2Ntar = rsMPC_S2Ntar
End Function

Function getRsMPC_S2Cijena() As Integer
    getRsMPC_S2Cijena = rsMPC_S2Cijena
End Function

Function getRsMPC_S2Datum() As Integer
    getRsMPC_S2Datum = rsMPC_S2Datum
End Function

Function getRsMPC_S3Ntar() As Integer
    getRsMPC_S3Ntar = rsMPC_S3Ntar
End Function

Function getRsMPC_S3Cijena() As Integer
    getRsMPC_S3Cijena = rsMPC_S3Cijena
End Function

Function getRsMPC_S3Datum() As Integer
    getRsMPC_S3Datum = rsMPC_S3Datum
End Function

Function getRsMPC_KAMPNtar() As Integer
    getRsMPC_KAMPNtar = rsMPC_KAMPNtar
End Function

Function getRsMPC_KAMPCijena() As Integer
    getRsMPC_KAMPCijena = rsMPC_KAMPCijena
End Function

Function getRsMPC_KAMPDatum() As Integer
    getRsMPC_KAMPDatum = rsMPC_KAMPDatum
End Function

Function getRsPoreznaGrupa() As Integer
    getRsPoreznaGrupa = rsPoreznaGrupa
End Function

Function getRsCEXV() As Integer
    getRsCEXV = rsCEXV
End Function


Function getColSifraArtikla() As String
    getColSifraArtikla = colSifraArtikla
End Function

Function getColBarkodArtikla() As String
    getColBarkodArtikla = colBarkodArtikla
End Function

Function getColNazivArtikla() As String
    getColNazivArtikla = colNazivArtikla
End Function

Function getColBrand() As String
    getColBrand = colBrand
End Function

Function getColPrincipal() As String
    getColPrincipal = colPrincipal
End Function

Function getColNivo1() As String
    getColNivo1 = colNivo1
End Function

Function getColNaziv1() As String
    getColNaziv1 = colNaziv1
End Function

Function getColNivo2() As String
    getColNivo2 = colNivo2
End Function

Function getColNaziv2() As String
    getColNaziv2 = colNaziv2
End Function

Function getColNivo3() As String
    getColNivo3 = colNivo3
End Function

Function getColNaziv3() As String
    getColNaziv3 = colNaziv3
End Function

Function getColNivo4() As String
    getColNivo4 = colNivo4
End Function

Function getColNaziv4() As String
    getColNaziv4 = colNaziv4
End Function

Function getColNivo5() As String
    getColNivo5 = colNivo5
End Function

Function getColNaziv5() As String
    getColNaziv5 = colNaziv5
End Function

Function getColAsortiman() As String
    getColAsortiman = colAsortiman
End Function

Function getColTSC() As String
    getColTSC = colTSC
End Function

Function getColOpis() As String
    getColOpis = colOpis
End Function

Function getColSvojstva() As String
    getColSvojstva = colSvojstva
End Function

Function getColPocetnaCijena() As String
    getColPocetnaCijena = colPocetnaCijena
End Function

Function getColMPC_ADatum() As String
    getColMPC_ADatum = colMPC_ADatum
End Function

Function getColMPC_ACijena() As String
    getColMPC_ACijena = colMPC_ACijena
End Function

Function getColMPC_ANovaCijena() As String
    getColMPC_ANovaCijena = colMPC_ANovaCijena
End Function

Function getColMPC_AIndeks() As String
    getColMPC_AIndeks = colMPC_AIndeks
End Function

Function getColMPC_BDatum() As String
    getColMPC_BDatum = colMPC_BDatum
End Function

Function getColMPC_BCijena() As String
    getColMPC_BCijena = colMPC_BCijena
End Function

Function getColMPC_BNovaCijena() As String
    getColMPC_BNovaCijena = colMPC_BNovaCijena
End Function

Function getColMPC_BIndeks() As String
    getColMPC_BIndeks = colMPC_BIndeks
End Function

Function getColMPC_CDatum() As String
    getColMPC_CDatum = colMPC_CDatum
End Function

Function getColMPC_CCijena() As String
    getColMPC_CCijena = colMPC_CCijena
End Function

Function getColMPC_CNovaCijena() As String
    getColMPC_CNovaCijena = colMPC_CNovaCijena
End Function

Function getColMPC_CIndeks() As String
    getColMPC_CIndeks = colMPC_CIndeks
End Function

Function getColMPC_DDatum() As String
    getColMPC_DDatum = colMPC_DDatum
End Function

Function getColMPC_DCijena() As String
    getColMPC_DCijena = colMPC_DCijena
End Function

Function getColMPC_DNovaCijena() As String
    getColMPC_DNovaCijena = colMPC_DNovaCijena
End Function

Function getColMPC_DIndeks() As String
    getColMPC_DIndeks = colMPC_DIndeks
End Function

Function getColMPC_S1Datum() As String
    getColMPC_S1Datum = colMPC_S1Datum
End Function

Function getColMPC_S1Cijena() As String
    getColMPC_S1Cijena = colMPC_S1Cijena
End Function

Function getColMPC_S1NovaCijena() As String
    getColMPC_S1NovaCijena = colMPC_S1NovaCijena
End Function

Function getColMPC_S1Indeks() As String
    getColMPC_S1Indeks = colMPC_S1Indeks
End Function

Function getColMPC_S2Datum() As String
    getColMPC_S2Datum = colMPC_S2Datum
End Function

Function getColMPC_S2Cijena() As String
    getColMPC_S2Cijena = colMPC_S2Cijena
End Function

Function getColMPC_S2NovaCijena() As String
    getColMPC_S2NovaCijena = colMPC_S2NovaCijena
End Function

Function getColMPC_S2Indeks() As String
    getColMPC_S2Indeks = colMPC_S2Indeks
End Function

Function getColMPC_S3Datum() As String
    getColMPC_S3Datum = colMPC_S3Datum
End Function

Function getColMPC_S3Cijena() As String
    getColMPC_S3Cijena = colMPC_S3Cijena
End Function

Function getColMPC_S3NovaCijena() As String
    getColMPC_S3NovaCijena = colMPC_S3NovaCijena
End Function

Function getColMPC_S3Indeks() As String
    getColMPC_S3Indeks = colMPC_S3Indeks
End Function

Function getColMPC_KAMPDatum() As String
    getColMPC_KAMPDatum = colMPC_KAMPDatum
End Function

Function getColMPC_KAMPCijena() As String
    getColMPC_KAMPCijena = colMPC_KAMPCijena
End Function

Function getColMPC_KAMPNovaCijena() As String
    getColMPC_KAMPNovaCijena = colMPC_KAMPNovaCijena
End Function

Function getColMPC_KAMPIndeks() As String
    getColMPC_KAMPIndeks = colMPC_KAMPIndeks
End Function


Function getColRedak() As String
    getColRedak = colRedak
End Function

Function getColPoreznaGrupa() As String
    getColPoreznaGrupa = colPoreznaGrupa
End Function

Function getColCEXV() As String
    getColCEXV = colCEXV
End Function

Function getColBrojPromjena() As String
    getColBrojPromjena = colBrojPromjena
End Function

