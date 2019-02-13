#Upisati u datoteku proizvoljan broj riječi.
#Korisnik nakon svakog unosa odlučuje želi li nastaviti unos.
#Ispisati iz datoteke samo one riječi kojima je
#broj pojavljivanja samoglasnika prost broji broj samoglasnika.

dat=open("datoteka.txt","w")
while True:
    rijec = input("Unesi rijec: ")
    dat.write(rijec+"\n")
    odg=input("Unesi 0 za kraj ili pritisni enter za nastavak: ")
    if odg == "0":
        break
    
dat.close()

dat=open("datoteka.txt","r")
rijeci=dat.readlines()
dat.close()

for rijec in rijeci:
    rijec = rijec[:-1]
    brojac_samoglasnika = 0
    for slovo in rijec:
        if slovo.upper() == "A" or slovo.upper() == "E" or slovo.upper() == "I" or slovo.upper() == "O" or slovo.upper() == "U":
            brojac_samoglasnika += 1

    if (brojac_samoglasnika > 1):
        prost = True
        for i in range(2, brojac_samoglasnika):
            if brojac_samoglasnika % i == 0:
                prost = False

        if not (prost):
            continue
        
        print("Riječ:", rijec, "- broj samoglasnika:", brojac_samoglasnika)
