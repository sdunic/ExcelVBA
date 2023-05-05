Attribute VB_Name = "cfg"
'config dokument u kojemu se:
' - inicijalno generira db connection string
' - parametriziraju stupci u tablici
' - parametriziraju indeksi recordseta kojeg uèitavamo iz Oracle baze podataka
Dim colSifraArtikla, colBarkodArtikla, colNazivArtikla, colBrand, colPrincipal As String
Dim colNivo1, colNaziv1, colNivo2, colNaziv2, colNivo3, colNaziv3, colNivo4, colNaziv4, colNivo5, colNaziv5 As String
Dim colAsortiman, colTSC, colOpis, colSvojstva As String
Dim colNA_Datum, colNA_Cijena, colNA_NovaCijena, colNA_Indeks, colIA_Datum, colIA_Cijena, colIA_NovaCijena, colIA_Indeks, colKatalog_Datum, colKatalog_Cijena, colKatalog_NovaCijena, colKatalog_Indeks, colRasprodaja_Datum, colRasprodaja_Cijena, colRasprodaja_NovaCijena, colRasprodaja_Indeks, colIstekRoka_Datum, colIstekRoka_Cijena, colIstekRoka_NovaCijena, colIstekRoka_Indeks As String
Dim colRedak, colPoreznaGrupa, colCEXV, colBrojPromjena As String

Dim rsSifraArtikla, rsBarkodArtikla, rsNazivArtikla, rsBrand, rsPrincipal As Integer
Dim rsNivo1, rsNaziv1, rsNivo2, rsNaziv2, rsNivo3, rsNaziv3, rsNivo4, rsNaziv4, rsNivo5, rsNaziv5 As Integer
Dim rsAsortiman, rsTSC, rsOpis, rsSvojstva As Integer
Dim rsNA_Ntar, rsNA_Cijena, rsNA_Datum, rsIA_Ntar, rsIA_Cijena, rsIA_Datum, rsKatalog_Ntar, rsKatalog_Cijena, rsKatalog_Datum, rsRasprodaja_Ntar, rsRasprodaja_Cijena, rsRasprodaja_Datum, rsIstekRoka_Ntar, rsIstekRoka_Cijena, rsIstekRoka_Datum As Integer
Dim rsPoreznaGrupa, rsCEXV As Integer

Dim rsNA_DatumKraja, rsIA_DatumKraja, rsKatalog_DatumKraja, rsRasprodaja_DatumKraja, rsIstekRoka_DatumKraja As Integer


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
    colNA_Datum = "U"
    colNA_Cijena = "V"
    colNA_NovaCijena = "W"
    colNA_Indeks = "X"
    colIA_Datum = "Y"
    colIA_Cijena = "Z"
    colIA_NovaCijena = "AA"
    colIA_Indeks = "AB"
    colKatalog_Datum = "AC"
    colKatalog_Cijena = "AD"
    colKatalog_NovaCijena = "AE"
    colKatalog_Indeks = "AF"
    colRasprodaja_Datum = "AG"
    colRasprodaja_Cijena = "AH"
    colRasprodaja_NovaCijena = "AI"
    colRasprodaja_Indeks = "AJ"
    colIstekRoka_Datum = "AK"
    colIstekRoka_Cijena = "AL"
    colIstekRoka_NovaCijena = "AM"
    colIstekRoka_Indeks = "AN"
    colRedak = "AO"
    colPoreznaGrupa = "AP"
    colCEXV = "AQ"
    colBrojPromjena = "AR"


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
    
    rsNA_Ntar = 19
    rsNA_Cijena = 20
    rsNA_Datum = 21
    rsNA_DatumKraja = 22
    
    rsIA_Ntar = 23
    rsIA_Cijena = 24
    rsIA_Datum = 25
    rsIA_DatumKraja = 26
    
    rsKatalog_Ntar = 27
    rsKatalog_Cijena = 28
    rsKatalog_Datum = 29
    rsKatalog_DatumKraja = 30
    
    rsRasprodaja_Ntar = 31
    rsRasprodaja_Cijena = 32
    rsRasprodaja_Datum = 33
    rsRasprodaja_DatumKraja = 34
    
    rsIstekRoka_Ntar = 35
    rsIstekRoka_Cijena = 36
    rsIstekRoka_Datum = 37
    rsIstekRoka_DatumKraja = 38
    
    rsPoreznaGrupa = 39
    rsCEXV = 40

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

Function getRsNA_Ntar() As Integer
    getRsNA_Ntar = rsNA_Ntar
End Function

Function getRsNA_Cijena() As Integer
    getRsNA_Cijena = rsNA_Cijena
End Function

Function getRsNA_Datum() As Integer
    getRsNA_Datum = rsNA_Datum
End Function
Function getRsNA_DatumKraja() As Integer
    getRsNA_DatumKraja = rsNA_DatumKraja
End Function

Function getRsIA_Ntar() As Integer
    getRsIA_Ntar = rsIA_Ntar
End Function

Function getRsIA_Cijena() As Integer
    getRsIA_Cijena = rsIA_Cijena
End Function

Function getRsIA_Datum() As Integer
    getRsIA_Datum = rsIA_Datum
End Function
Function getRsIA_DatumKraja() As Integer
    getRsIA_DatumKraja = rsIA_DatumKraja
End Function

Function getRsKatalog_Ntar() As Integer
    getRsKatalog_Ntar = rsKatalog_Ntar
End Function

Function getRsKatalog_Cijena() As Integer
    getRsKatalog_Cijena = rsKatalog_Cijena
End Function

Function getRsKatalog_Datum() As Integer
    getRsKatalog_Datum = rsKatalog_Datum
End Function
Function getRsKatalog_DatumKraja() As Integer
    getRsKatalog_DatumKraja = rsKatalog_DatumKraja
End Function

Function getRsRasprodaja_Ntar() As Integer
    getRsRasprodaja_Ntar = rsRasprodaja_Ntar
End Function

Function getRsRasprodaja_Cijena() As Integer
    getRsRasprodaja_Cijena = rsRasprodaja_Cijena
End Function

Function getRsRasprodaja_Datum() As Integer
    getRsRasprodaja_Datum = rsRasprodaja_Datum
End Function
Function getRsRasprodaja_DatumKraja() As Integer
    getRsRasprodaja_DatumKraja = rsRasprodaja_DatumKraja
End Function

Function getRsIstekRoka_Ntar() As Integer
    getRsIstekRoka_Ntar = rsIstekRoka_Ntar
End Function

Function getRsIstekRoka_Cijena() As Integer
    getRsIstekRoka_Cijena = rsIstekRoka_Cijena
End Function

Function getRsIstekRoka_Datum() As Integer
    getRsIstekRoka_Datum = rsIstekRoka_Datum
End Function
Function getRsIstekRoka_DatumKraja() As Integer
    getRsIstekRoka_DatumKraja = rsIstekRoka_DatumKraja
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

Function getColNA_Datum() As String
    getColNA_Datum = colNA_Datum
End Function

Function getColNA_Cijena() As String
    getColNA_Cijena = colNA_Cijena
End Function

Function getColNA_NovaCijena() As String
    getColNA_NovaCijena = colNA_NovaCijena
End Function

Function getColNA_Indeks() As String
    getColNA_Indeks = colNA_Indeks
End Function

Function getColIA_Datum() As String
    getColIA_Datum = colIA_Datum
End Function

Function getColIA_Cijena() As String
    getColIA_Cijena = colIA_Cijena
End Function

Function getColIA_NovaCijena() As String
    getColIA_NovaCijena = colIA_NovaCijena
End Function

Function getColIA_Indeks() As String
    getColIA_Indeks = colIA_Indeks
End Function

Function getColKatalog_Datum() As String
    getColKatalog_Datum = colKatalog_Datum
End Function

Function getColKatalog_Cijena() As String
    getColKatalog_Cijena = colKatalog_Cijena
End Function

Function getColKatalog_NovaCijena() As String
    getColKatalog_NovaCijena = colKatalog_NovaCijena
End Function

Function getColKatalog_Indeks() As String
    getColKatalog_Indeks = colKatalog_Indeks
End Function

Function getColRasprodaja_Datum() As String
    getColRasprodaja_Datum = colRasprodaja_Datum
End Function

Function getColRasprodaja_Cijena() As String
    getColRasprodaja_Cijena = colRasprodaja_Cijena
End Function

Function getColRasprodaja_NovaCijena() As String
    getColRasprodaja_NovaCijena = colRasprodaja_NovaCijena
End Function

Function getColRasprodaja_Indeks() As String
    getColRasprodaja_Indeks = colRasprodaja_Indeks
End Function

Function getColIstekRoka_Datum() As String
    getColIstekRoka_Datum = colIstekRoka_Datum
End Function

Function getColIstekRoka_Cijena() As String
    getColIstekRoka_Cijena = colIstekRoka_Cijena
End Function

Function getColIstekRoka_NovaCijena() As String
    getColIstekRoka_NovaCijena = colIstekRoka_NovaCijena
End Function

Function getColIstekRoka_Indeks() As String
    getColIstekRoka_Indeks = colIstekRoka_Indeks
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


