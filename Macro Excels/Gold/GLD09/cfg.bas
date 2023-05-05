Attribute VB_Name = "cfg"
Dim korisnik, lokacija, tipFakture, kupac, ugovor, datumFakture, artikl, lv_lu, kolicina, ukupniIznos, tm, robniCvor, analitickiArtikl, napomena As String
Dim analitickiTM, analitickiMrezniCvor As String
Dim reasonCodeTekst As String
Dim reasonCodeRedak As Integer

Dim zaglavlje, stavke As Integer

Sub Init()
    'parametrizacija stupaca
    zaglavlje = 5
    korisnik = "B"
    lokacija = "C"
    tipFakture = "D"
    kupac = "E"
    ugovor = "F"
    datumFakture = "G"
    napomena = "H"
    
    reasonCodeTekst = "C"
    reasonCodeRedak = 2
        
    stavke = 11
    artikl = "B"
    lv_lu = "C"
    kolicina = "D"
    ukupniIznos = "E"
    tm = "F"
    robniCvor = "H"
    analitickiArtikl = "G"
    analitickiTM = "I"
    analitickiMrezniCvor = "J"
    
        
End Sub

Function get_zaglavlje() As Integer
    get_zaglavlje = zaglavlje
End Function

Function get_stavke() As Integer
    get_stavke = stavke
End Function

Function get_korisnik() As String
    get_korisnik = korisnik
End Function

Function get_lokacija() As String
    get_lokacija = lokacija
End Function

Function get_tipFakture() As String
    get_tipFakture = tipFakture
End Function

Function get_kupac() As String
    get_kupac = kupac
End Function

Function get_ugovor() As String
    get_ugovor = ugovor
End Function

Function get_datumFakture() As String
    get_datumFakture = datumFakture
End Function

Function get_napomena() As String
    get_napomena = napomena
End Function

Function get_artikl() As String
    get_artikl = artikl
End Function

Function get_kolicina() As String
    get_kolicina = kolicina
End Function

Function get_lv_lu() As String
    get_lv_lu = lv_lu
End Function

Function get_ukupniIznos() As String
    get_ukupniIznos = ukupniIznos
End Function

Function get_tm() As String
    get_tm = tm
End Function

Function get_robniCvor() As String
    get_robniCvor = robniCvor
End Function

Function get_analitickiArtikl() As String
    get_analitickiArtikl = analitickiArtikl
End Function

Function get_analitickiTM() As String
    get_analitickiTM = analitickiTM
End Function

Function get_analitickiMrezniCvor() As String
    get_analitickiMrezniCvor = analitickiMrezniCvor
End Function

Function get_reasonCodeTekst() As String
    get_reasonCodeTekst = reasonCodeTekst
End Function

Function get_reasonCodeRedak() As Integer
    get_reasonCodeRedak = reasonCodeRedak
End Function

