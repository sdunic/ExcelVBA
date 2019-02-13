#Upisati u datoteku proizvoljan broj riječi. Korisnik nakon unosa svake riječi odlučuje želi li nastaviti unos. Upisati jedno slovo.
#Ispisati iz datoteke samo one riječi u kojima je broj pojavljivanja unesenog slova djeljiv sa 3, kao i broj pojavljivanja tog slova.

def provjera(rijec, slovo):
    brojac = 0
    for s in rijec:
        if s == slovo:
            brojac+=1
    if brojac % 3 == 0 and brojac > 0:
        return brojac
    return 0
    
dat=open("datoteka.txt", "w")

while True:
    rijec=input("Unesite riječ: ")
    dat.write(rijec + "\n")

    odg = input("Želite li nastaviti unos? (DA/NE)")
    if(odg.upper()=="NE"):
        break
dat.close()


slovo = str(input("Unesi slovo: "))

dat=open("datoteka.txt", "r")
for rijec in dat:
    rezultat = provjera(rijec[:-1], slovo)
    if rezultat > 0:
       print(rijec[:-1], "- slovo", slovo, "se pojavljuje", rezultat, "puta")
    
        

