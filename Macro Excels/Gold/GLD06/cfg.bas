Attribute VB_Name = "cfg"
'config dokument u kojemu se:
' - inicijalno generira db connection string
' - parametriziraju stupci u tablici
' - parametriziraju indeksi recordseta kojeg uèitavamo iz Oracle baze podataka
Dim colSifraArtikla, colBarkodArtikla, colNazivArtikla, colBrand, colPrincipal As String
Dim colNivo1, colNaziv1, colNivo2, colNaziv2, colNivo3, colNaziv3, colNivo4, colNaziv4, colNivo5, colNaziv5 As String
Dim colAsortiman, colTSC, colOpis, colSvojstva As String
Dim colKonzumHiperDatum, colKonzumHiperCijena, colKonzumHiperNovaCijena, colKonzumHiperIndeks, colKonzumMaxiDatum, colKonzumMaxiCijena, colKonzumMaxiNovaCijena, colKonzumMaxiIndeks, colStudenacDatum, colStudenacCijena, colStudenacNovaCijena, colStudenacIndeks
Dim colRedak, colPoreznaGrupa, colCEXV, colBrojPromjena As String

Dim rsSifraArtikla, rsBarkodArtikla, rsNazivArtikla, rsBrand, rsPrincipal As Integer
Dim rsNivo1, rsNaziv1, rsNivo2, rsNaziv2, rsNivo3, rsNaziv3, rsNivo4, rsNaziv4, rsNivo5, rsNaziv5 As Integer
Dim rsAsortiman, rsTSC, rsOpis, rsSvojstva As Integer
Dim rsKonzumHiperNtar, rsKonzumHiperCijena, rsKonzumHiperDatum, rsKonzumMaxiNtar, rsKonzumMaxiCijena, rsKonzumMaxiDatum, rsStudenacNtar, rsStudenacCijena, rsStudenacDatum
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
    colKonzumHiperDatum = "U"
    colKonzumHiperCijena = "V"
    colKonzumHiperNovaCijena = "W"
    colKonzumHiperIndeks = "X"
    colKonzumMaxiDatum = "Y"
    colKonzumMaxiCijena = "Z"
    colKonzumMaxiNovaCijena = "AA"
    colKonzumMaxiIndeks = "AB"
    colStudenacDatum = "AC"
    colStudenacCijena = "AD"
    colStudenacNovaCijena = "AE"
    colStudenacIndeks = "AF"
    colRedak = "AG"
    colPoreznaGrupa = "AH"
    colCEXV = "AI"
    colBrojPromjena = "AJ"


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
    rsKonzumHiperNtar = 19
    rsKonzumHiperCijena = 20
    rsKonzumHiperDatum = 21
    rsKonzumMaxiNtar = 22
    rsKonzumMaxiCijena = 23
    rsKonzumMaxiDatum = 24
    rsStudenacNtar = 25
    rsStudenacCijena = 26
    rsStudenacDatum = 27
    rsPoreznaGrupa = 28
    rsCEXV = 29

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

Function getRsKonzumHiperNtar() As Integer
    getRsKonzumHiperNtar = rsKonzumHiperNtar
End Function

Function getRsKonzumHiperCijena() As Integer
    getRsKonzumHiperCijena = rsKonzumHiperCijena
End Function

Function getRsKonzumHiperDatum() As Integer
    getRsKonzumHiperDatum = rsKonzumHiperDatum
End Function

Function getRsKonzumMaxiNtar() As Integer
    getRsKonzumMaxiNtar = rsKonzumMaxiNtar
End Function

Function getRsKonzumMaxiCijena() As Integer
    getRsKonzumMaxiCijena = rsKonzumMaxiCijena
End Function

Function getRsKonzumMaxiDatum() As Integer
    getRsKonzumMaxiDatum = rsKonzumMaxiDatum
End Function

Function getRsStudenacNtar() As Integer
    getRsStudenacNtar = rsStudenacNtar
End Function

Function getRsStudenacCijena() As Integer
    getRsStudenacCijena = rsStudenacCijena
End Function

Function getRsStudenacDatum() As Integer
    getRsStudenacDatum = rsStudenacDatum
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

Function getRsMPC_SNtar() As Integer
    getRsMPC_SNtar = rsMPC_SNtar
End Function

Function getRsMPC_SCijena() As Integer
    getRsMPC_SCijena = rsMPC_SCijena
End Function

Function getRsMPC_SDatum() As Integer
    getRsMPC_SDatum = rsMPC_SDatum
End Function

Function getRsTNC_Ntar() As Integer
    getRsTNC_Ntar = rsTNC_Ntar
End Function

Function getRsTNC_Cijena() As Integer
    getRsTNC_Cijena = rsTNC_Cijena
End Function

Function getRsTNC_Datum() As Integer
    getRsTNC_Datum = rsTNC_Datum
End Function

Function getRsTNC_ANtar() As Integer
    getRsTNC_ANtar = rsTNC_ANtar
End Function

Function getRsTNC_ACijena() As Integer
    getRsTNC_ACijena = rsTNC_ACijena
End Function

Function getRsTNC_ADatum() As Integer
    getRsTNC_ADatum = rsTNC_ADatum
End Function

Function getRsTNC_BNtar() As Integer
    getRsTNC_BNtar = rsTNC_BNtar
End Function

Function getRsTNC_BCijena() As Integer
    getRsTNC_BCijena = rsTNC_BCijena
End Function

Function getRsTNC_BDatum() As Integer
    getRsTNC_BDatum = rsTNC_BDatum
End Function

Function getRsTNC_CNtar() As Integer
    getRsTNC_CNtar = rsTNC_CNtar
End Function

Function getRsTNC_CCijena() As Integer
    getRsTNC_CCijena = rsTNC_CCijena
End Function

Function getRsTNC_CDatum() As Integer
    getRsTNC_CDatum = rsTNC_CDatum
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

Function getColKonzumHiperDatum() As String
    getColKonzumHiperDatum = colKonzumHiperDatum
End Function

Function getColKonzumHiperCijena() As String
    getColKonzumHiperCijena = colKonzumHiperCijena
End Function

Function getColKonzumHiperNovaCijena() As String
    getColKonzumHiperNovaCijena = colKonzumHiperNovaCijena
End Function

Function getColKonzumHiperIndeks() As String
    getColKonzumHiperIndeks = colKonzumHiperIndeks
End Function

Function getColKonzumMaxiDatum() As String
    getColKonzumMaxiDatum = colKonzumMaxiDatum
End Function

Function getColKonzumMaxiCijena() As String
    getColKonzumMaxiCijena = colKonzumMaxiCijena
End Function

Function getColKonzumMaxiNovaCijena() As String
    getColKonzumMaxiNovaCijena = colKonzumMaxiNovaCijena
End Function

Function getColKonzumMaxiIndeks() As String
    getColKonzumMaxiIndeks = colKonzumMaxiIndeks
End Function

Function getColStudenacDatum() As String
    getColStudenacDatum = colStudenacDatum
End Function

Function getColStudenacCijena() As String
    getColStudenacCijena = colStudenacCijena
End Function

Function getColStudenacNovaCijena() As String
    getColStudenacNovaCijena = colStudenacNovaCijena
End Function

Function getColStudenacIndeks() As String
    getColStudenacIndeks = colStudenacIndeks
End Function

Function getColMPC_DDatum() As String
    getColMPC_DDatum = colMPC_DDatum
End Function

Function getColMPC_DCijena() As String
    getColMPC_DCijena = colMPC_DCijena
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

