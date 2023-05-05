Attribute VB_Name = "cfg"
'config dokument u kojemu se:
' - inicijalno generira db connection string
' - parametriziraju stupci u tablici
' - parametriziraju indeksi recordseta kojeg uèitavamo iz Oracle baze podataka
Dim colSifraArtikla, colBarkodArtikla, colNazivArtikla, colBrand, colPrincipal As String
Dim colNivo1, colNaziv1, colNivo2, colNaziv2, colNivo3, colNaziv3, colNivo4, colNaziv4, colNivo5, colNaziv5 As String
Dim colAsortiman, colTSC, colOpis, colSvojstva, colPocetnaCijena As String
Dim colMPC_KAMPDatum, colMPC_KAMPCijena, colMPC_KAMPNovaCijena, colMPC_KAMPIndeks, colTNC_KAMPDatum, colTNC_KAMPCijena, colTNC_KAMPNovaCijena, colTNC_KAMPIndeks As String
Dim colRedak, colPoreznaGrupa, colCEXV, colBrojPromjena As String

Dim rsSifraArtikla, rsBarkodArtikla, rsNazivArtikla, rsBrand, rsPrincipal As Integer
Dim rsNivo1, rsNaziv1, rsNivo2, rsNaziv2, rsNivo3, rsNaziv3, rsNivo4, rsNaziv4, rsNivo5, rsNaziv5 As Integer
Dim rsAsortiman, rsTSC, rsOpis, rsSvojstva As Integer
Dim rsMPC_KAMPNtar, rsMPC_KAMPCijena, rsMPC_KAMPDatum, rsTNC_KAMPNtar, rsTNC_KAMPCijena, rsTNC_KAMPDatum As Integer
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
    colMPC_KAMPDatum = "V"
    colMPC_KAMPCijena = "W"
    colMPC_KAMPNovaCijena = "X"
    colMPC_KAMPIndeks = "Y"
    colTNC_KAMPDatum = "Z"
    colTNC_KAMPCijena = "AA"
    colTNC_KAMPNovaCijena = "AB"
    colTNC_KAMPIndeks = "AC"
    colRedak = "AD"
    colPoreznaGrupa = "AE"
    colCEXV = "AF"
    colBrojPromjena = "AG"

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
    rsMPC_KAMPNtar = 19
    rsMPC_KAMPCijena = 20
    rsMPC_KAMPDatum = 21
    rsTNC_KAMPNtar = 22
    rsTNC_KAMPCijena = 23
    rsTNC_KAMPDatum = 24
    rsPoreznaGrupa = 25
    rsCEXV = 26

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

Function getRsMPC_KAMPNtar() As Integer
    getRsMPC_KAMPNtar = rsMPC_KAMPNtar
End Function

Function getRsMPC_KAMPCijena() As Integer
    getRsMPC_KAMPCijena = rsMPC_KAMPCijena
End Function

Function getRsMPC_KAMPDatum() As Integer
    getRsMPC_KAMPDatum = rsMPC_KAMPDatum
End Function

Function getRsTNC_KAMPNtar() As Integer
    getRsTNC_KAMPNtar = rsTNC_KAMPNtar
End Function

Function getRsTNC_KAMPCijena() As Integer
    getRsTNC_KAMPCijena = rsTNC_KAMPCijena
End Function

Function getRsTNC_KAMPDatum() As Integer
    getRsTNC_KAMPDatum = rsTNC_KAMPDatum
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

Function getColTNC_KAMPDatum() As String
    getColTNC_KAMPDatum = colTNC_KAMPDatum
End Function

Function getColTNC_KAMPCijena() As String
    getColTNC_KAMPCijena = colTNC_KAMPCijena
End Function

Function getColTNC_KAMPNovaCijena() As String
    getColTNC_KAMPNovaCijena = colTNC_KAMPNovaCijena
End Function

Function getColTNC_KAMPIndeks() As String
    getColTNC_KAMPIndeks = colTNC_KAMPIndeks
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



