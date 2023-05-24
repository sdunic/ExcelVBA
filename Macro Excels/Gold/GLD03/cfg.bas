Attribute VB_Name = "cfg"
'config dokument u kojemu se:
' - inicijalno generira db connection string
' - parametriziraju stupci u tablici
' - parametriziraju indeksi recordseta kojeg u�itavamo iz Oracle baze podataka
Dim colSifraArtikla, colBarkodArtikla, colNazivArtikla, colBrand, colPrincipal As String
Dim colNivo1, colNaziv1, colNivo2, colNaziv2, colNivo3, colNaziv3, colNivo4, colNaziv4, colNivo5, colNaziv5 As String
Dim colAsortiman, colTSC, colOpis, colSvojstva As String
Dim colTNC_ADatum, colTNC_ACijena, colTNC_ANovaCijena, colTNC_AIndeks, colTNC_BDatum, colTNC_BCijena, colTNC_BNovaCijena, colTNC_BIndeks, colTNC_CDatum, colTNC_CCijena, colTNC_CNovaCijena, colTNC_CIndeks, colTNC_DDatum, colTNC_DCijena, colTNC_DNovaCijena, colTNC_DIndeks, colTNC_SDatum, colTNC_SCijena, colTNC_SNovaCijena, colTNC_SIndeks, colTNC_KAMPDatum, colTNC_KAMPCijena, colTNC_KAMPNovaCijena, colTNC_KAMPIndeks As String
Dim colTNC_Datum, colTNC_Cijena, colTNC_NovaCijena, colTNC_Indeks As String
Dim colRedak, colPoreznaGrupa, colCEXV, colBrojPromjena As String

Dim rsSifraArtikla, rsBarkodArtikla, rsNazivArtikla, rsBrand, rsPrincipal As Integer
Dim rsNivo1, rsNaziv1, rsNivo2, rsNaziv2, rsNivo3, rsNaziv3, rsNivo4, rsNaziv4, rsNivo5, rsNaziv5 As Integer
Dim rsAsortiman, rsTSC, rsOpis, rsSvojstva As Integer
Dim rsTNC_ANtar, rsTNC_ACijena, rsTNC_ADatum, rsTNC_BNtar, rsTNC_BCijena, rsTNC_BDatum, rsTNC_CNtar, rsTNC_CCijena, rsTNC_CDatum, rsTNC_DNtar, rsTNC_DCijena, rsTNC_DDatum, rsTNC_SNtar, rsTNC_SCijena, rsTNC_SDatum, rsTNC_KAMPNtar, rsTNC_KAMPCijena, rsTNC_KAMPDatum As Integer
Dim rsTNC_Ntar, rsTNC_Cijena, rsTNC_Datum As Integer
Dim rsPoreznaGrupa, rsCEXV As Integer

Dim rsTNC_DatumKraja, rsTNC_ADatumKraja, rsTNC_BDatumKraja, rsTNC_CDatumKraja, rsTNC_DDatumKraja, rsTNC_SDatumKraja, rsTNC_KAMPDatumKraja As Integer


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
    colTNC_Datum = "U"
    colTNC_Cijena = "V"
    colTNC_NovaCijena = "W"
    colTNC_Indeks = "X"
    colTNC_ADatum = "Y"
    colTNC_ACijena = "Z"
    colTNC_ANovaCijena = "AA"
    colTNC_AIndeks = "AB"
    colTNC_BDatum = "AC"
    colTNC_BCijena = "AD"
    colTNC_BNovaCijena = "AE"
    colTNC_BIndeks = "AF"
    colTNC_CDatum = "AG"
    colTNC_CCijena = "AH"
    colTNC_CNovaCijena = "AI"
    colTNC_CIndeks = "AJ"
    colTNC_DDatum = "AK"
    colTNC_DCijena = "AL"
    colTNC_DNovaCijena = "AM"
    colTNC_DIndeks = "AN"
    colTNC_SDatum = "AO"
    colTNC_SCijena = "AP"
    colTNC_SNovaCijena = "AQ"
    colTNC_SIndeks = "AR"
    colTNC_KAMPDatum = "AS"
    colTNC_KAMPCijena = "AT"
    colTNC_KAMPNovaCijena = "AU"
    colTNC_KAMPIndeks = "AV"
    colRedak = "AW"
    colPoreznaGrupa = "AX"
    colCEXV = "AY"
    colBrojPromjena = "AZ"


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
    
    rsTNC_Ntar = 19
    rsTNC_Cijena = 20
    rsTNC_Datum = 21
    rsTNC_DatumKraja = 22
    
    rsTNC_ANtar = 23
    rsTNC_ACijena = 24
    rsTNC_ADatum = 25
    rsTNC_ADatumKraja = 26
    
    rsTNC_BNtar = 27
    rsTNC_BCijena = 28
    rsTNC_BDatum = 29
    rsTNC_BDatumKraja = 30
    
    rsTNC_CNtar = 31
    rsTNC_CCijena = 32
    rsTNC_CDatum = 33
    rsTNC_CDatumKraja = 34
    
    rsTNC_DNtar = 35
    rsTNC_DCijena = 36
    rsTNC_DDatum = 37
    rsTNC_DDatumKraja = 38
    
    rsTNC_SNtar = 39
    rsTNC_SCijena = 40
    rsTNC_SDatum = 41
    rsTNC_SDatumKraja = 42
    
    rsTNC_KAMPNtar = 43
    rsTNC_KAMPCijena = 44
    rsTNC_KAMPDatum = 45
    rsTNC_KAMPDatumKraja = 46
    
    rsPoreznaGrupa = 47
    rsCEXV = 48

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

Function getRsTNC_Ntar() As Integer
    getRsTNC_Ntar = rsTNC_Ntar
End Function

Function getRsTNC_Cijena() As Integer
    getRsTNC_Cijena = rsTNC_Cijena
End Function

Function getRsTNC_Datum() As Integer
    getRsTNC_Datum = rsTNC_Datum
End Function
Function getRsTNC_DatumKraja() As Integer
    getRsTNC_DatumKraja = rsTNC_DatumKraja
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
Function getRsTNC_ADatumKraja() As Integer
    getRsTNC_ADatumKraja = rsTNC_ADatumKraja
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
Function getRsTNC_BDatumKraja() As Integer
    getRsTNC_BDatumKraja = rsTNC_BDatumKraja
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
Function getRsTNC_CDatumKraja() As Integer
    getRsTNC_CDatumKraja = rsTNC_CDatumKraja
End Function

Function getRsTNC_DNtar() As Integer
    getRsTNC_DNtar = rsTNC_DNtar
End Function

Function getRsTNC_DCijena() As Integer
    getRsTNC_DCijena = rsTNC_DCijena
End Function

Function getRsTNC_DDatum() As Integer
    getRsTNC_DDatum = rsTNC_DDatum
End Function

Function getRsTNC_DDatumKraja() As Integer
    getRsTNC_DDatumKraja = rsTNC_DDatumKraja
End Function

Function getRsTNC_SNtar() As Integer
    getRsTNC_SNtar = rsTNC_SNtar
End Function

Function getRsTNC_SCijena() As Integer
    getRsTNC_SCijena = rsTNC_SCijena
End Function

Function getRsTNC_SDatum() As Integer
    getRsTNC_SDatum = rsTNC_SDatum
End Function
Function getRsTNC_SDatumKraja() As Integer
    getRsTNC_SDatumKraja = rsTNC_SDatumKraja
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
Function getRsTNC_KAMPDatumKraja() As Integer
    getRsTNC_KAMPDatumKraja = rsTNC_KAMPDatumKraja
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

Function getColTNC_Datum() As String
    getColTNC_Datum = colTNC_Datum
End Function

Function getColTNC_Cijena() As String
    getColTNC_Cijena = colTNC_Cijena
End Function

Function getColTNC_NovaCijena() As String
    getColTNC_NovaCijena = colTNC_NovaCijena
End Function

Function getColTNC_Indeks() As String
    getColTNC_Indeks = colTNC_Indeks
End Function

Function getColTNC_ADatum() As String
    getColTNC_ADatum = colTNC_ADatum
End Function

Function getColTNC_ACijena() As String
    getColTNC_ACijena = colTNC_ACijena
End Function

Function getColTNC_ANovaCijena() As String
    getColTNC_ANovaCijena = colTNC_ANovaCijena
End Function

Function getColTNC_AIndeks() As String
    getColTNC_AIndeks = colTNC_AIndeks
End Function

Function getColTNC_BDatum() As String
    getColTNC_BDatum = colTNC_BDatum
End Function

Function getColTNC_BCijena() As String
    getColTNC_BCijena = colTNC_BCijena
End Function

Function getColTNC_BNovaCijena() As String
    getColTNC_BNovaCijena = colTNC_BNovaCijena
End Function

Function getColTNC_BIndeks() As String
    getColTNC_BIndeks = colTNC_BIndeks
End Function

Function getColTNC_CDatum() As String
    getColTNC_CDatum = colTNC_CDatum
End Function

Function getColTNC_CCijena() As String
    getColTNC_CCijena = colTNC_CCijena
End Function

Function getColTNC_CNovaCijena() As String
    getColTNC_CNovaCijena = colTNC_CNovaCijena
End Function

Function getColTNC_CIndeks() As String
    getColTNC_CIndeks = colTNC_CIndeks
End Function

Function getColTNC_DDatum() As String
    getColTNC_DDatum = colTNC_DDatum
End Function

Function getColTNC_DCijena() As String
    getColTNC_DCijena = colTNC_DCijena
End Function

Function getColTNC_DNovaCijena() As String
    getColTNC_DNovaCijena = colTNC_DNovaCijena
End Function

Function getColTNC_DIndeks() As String
    getColTNC_DIndeks = colTNC_DIndeks
End Function

Function getColTNC_SDatum() As String
    getColTNC_SDatum = colTNC_SDatum
End Function

Function getColTNC_SCijena() As String
    getColTNC_SCijena = colTNC_SCijena
End Function

Function getColTNC_SNovaCijena() As String
    getColTNC_SNovaCijena = colTNC_SNovaCijena
End Function

Function getColTNC_SIndeks() As String
    getColTNC_SIndeks = colTNC_SIndeks
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

