Attribute VB_Name = "cfg"
'config dokument u kojemu se:
' - inicijalno generira db connection string
' - parametriziraju stupci u tablici
' - parametriziraju indeksi recordseta kojeg uèitavamo iz Oracle baze podataka

Dim colSifraArtikla, colBarkodArtikla, colNazivArtikla As String
Dim colBrand, colPrincipal As String
Dim colNivo1, colNaziv1, colNivo2, colNaziv2, colNivo3, colNaziv3, colNivo4, colNaziv4, colNivo5, colNaziv5 As String
Dim colTSC, colOpis, colSvojstva As String
Dim colNovaCijena As String
Dim colPoreznaGrupa, colCEXV As String
Dim colRedak As String
Dim colNTAR, colDdeb, colDfin, colPrix As String

Dim rsSifraArtikla, rsBarkodArtikla, rsNazivArtikla As Integer
Dim rsBrand, rsPrincipal As Integer
Dim rsNivo1, rsNaziv1, rsNivo2, rsNaziv2, rsNivo3, rsNaziv3, rsNivo4, rsNaziv4, rsNivo5, rsNaziv5 As Integer
Dim rsTSC, rsOpis, rsSvojstva As Integer
Dim rsDatumCijene, rsDatumKrajaCijene, rsCijena As Integer
Dim rsPoreznaGrupa, rsCEXV As Integer
Dim rsNTAR As Integer

Sub Init()
    
    'parametrizacija stupaca
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
    colTSC = "Q"
    colOpis = "R"
    colSvojstva = "S"
    colNTAR = "T"
    colDdeb = "U"
    colDfin = "V"
    colPrix = "W"
    colNovaCijena = "X"
    colRedak = "Y"
    colPoreznaGrupa = "Z"
    colCEXV = "AA"

    'parametrizacija poretka u recordsetu
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
    rsTSC = 15
    rsOpis = 16
    rsSvojstva = 17
    rsDatumCijene = 18
    rsDatumKrajaCijene = 19
    rsCijena = 20
    rsPoreznaGrupa = 21
    rsCEXV = 22
    rsNTAR = 23
    
End Sub

Function getColNTAR() As String
    getColNTAR = colNTAR
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

Function getColTSC() As String
    getColTSC = colTSC
End Function

Function getColOpis() As String
    getColOpis = colOpis
End Function

Function getColSvojstva() As String
    getColSvojstva = colSvojstva
End Function

Function getColDdeb() As String
    getColDdeb = colDdeb
End Function

Function getColDfin() As String
    getColDfin = colDfin
End Function

Function getColPrix() As String
    getColPrix = colPrix
End Function


Function getColNovaCijena() As String
    getColNovaCijena = colNovaCijena
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

Function getRsTSC() As Integer
    getRsTSC = rsTSC
End Function

Function getRsOpis() As Integer
    getRsOpis = rsOpis
End Function

Function getRsSvojstva() As Integer
    getRsSvojstva = rsSvojstva
End Function

Function getRsDatumCijene() As Integer
    getRsDatumCijene = rsDatumCijene
End Function
Function getRsDatumKrajaCijene() As Integer
    getRsDatumKrajaCijene = rsDatumKrajaCijene
End Function

Function getRsCijena() As Integer
    getRsCijena = rsCijena
End Function

Function getRsPoreznaGrupa() As Integer
    getRsPoreznaGrupa = rsPoreznaGrupa
End Function

Function getRsCEXV() As Integer
    getRsCEXV = rsCEXV
End Function

Function getRsNTAR() As Integer
    getRsNTAR = rsNTAR
End Function



