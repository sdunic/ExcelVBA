#Korisnik unosi rečenice u datoteku. Korisnik nakon svakog unosa odlučuje želi li nastaviti unos.
#Za svaku rečenicu treba ispisati svaku riječ zasebno jednu ispod druge te ispisati njihove duljine.

def razdvoji_rijeci(recenica):
    rijeci = []
    rijec = ""
    
    for i in range(len(recenica)):
        rijec += recenica[i]
        if (recenica[i] == " "):
            rijeci.append(rijec.strip())
            rijec = ""

        if (i == len(recenica) - 1 and recenica[:-1] != " "):
            rijeci.append(rijec)
    
    return rijeci
    

dat=open("datoteka.txt","w")
while True:
    recenica = input("Unesi recenicu: ")
    dat.write(recenica+"\n")
    odg=input("Unesi 0 za kraj ili pritisni enter za nastavak: ")
    if odg == "0":
        break
    
dat.close()

dat=open("datoteka.txt","r")
recenice=dat.readlines()
dat.close()

for recenica in recenice:
    recenica = recenica[:-1]
    print(recenica)

    for rijec in razdvoji_rijeci(recenica):
        print("Riječ:", rijec, "- duljina:", len(rijec))
