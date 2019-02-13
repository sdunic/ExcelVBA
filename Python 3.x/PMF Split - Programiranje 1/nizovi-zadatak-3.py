#Korisnik unosi prirodne brojeve u niz.
#Prilikom unosa izbrojati koliko je pogrešno unesenih.
#Korisnik prestaje unositi brojeve kad se unese 0.
#Program treba za svaki element niza ispisati zbroj prve i zadnje znamenke ako je paran,
#a ako je neparan treba ispisati njihov umnožak.


def izracunaj_prvu_znamenku(x):
    while x != 0:
        x = x // 10
        if x % 10 == x:
            return x

#definiramo početne uvjete
niz = []
pogresni_brojevi = 0

#unos u niz i provjera pogresnih unosa
#dovoljno provjeriti jesu li brojevi manji od nula jer je preduvjet da korisnik unosi prirodne brojeve
#break ako korisnik unese nulu
while True:
    broj = int(input("Unesi prirodan broj: "))
    
    if(broj == 0):
        break

    if(broj < 0):
        pogresni_brojevi += 1
    else:         
        niz.append(broj)

#for petlja kroz niz i provjerimo elemente i ispišemo rješenje
for broj in niz:
    prva_znamenka = izracunaj_prvu_znamenku(broj)
    #prva_znamenka = int(str(broj)[0]) - izračun prve znamenke preko stringa
    zadnja_znamenka = broj % 10
    if (broj % 2 == 0):
        #ako je paran
        print("Broj:", broj, "- rezultat:", prva_znamenka + zadnja_znamenka)
    else:
        #ako je neparan
        print("Broj:", broj, "- rezultat:", prva_znamenka * zadnja_znamenka)
        
