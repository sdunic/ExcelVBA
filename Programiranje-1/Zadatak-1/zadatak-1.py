#Prva tocka zadatka.
dat = open("zadatak-1.txt", "w")

ponovni_unos = ""

while ponovni_unos != "NE" :
    n = int(input("Unesi broj: "))
    if(n > 0):
        dat.write(str(n) + "\n")
    ponovni_unos = (input("Unosimo ponovo DA/NE - ")).upper()

dat.close()


#Druga tocka zadatka.
dat = open("zadatak-1.txt", "r")

n_total = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0] #Treca tocka zadatka.

for redak in dat:
    n = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
    for znak in redak:
        if(znak != "\n"):
            n[int(znak)]+=1
            n_total[int(znak)]+=1 #Treca tocka zadatka.

    redak_za_ispis = str(int(redak))+ ":\t\t"
    for i in range(len(n)):
        redak_za_ispis += (str(i)+"-"+str(n[i]))
        if(i != len(n)-1):
            redak_za_ispis += ", "
            
    print(redak_za_ispis)
dat.close()


#Cetvrta tocka zadatka.
kolicina_najvise_pojavljivanja = -1
znamenka_najvise_pojavljivanja = -1

redak_za_ispis = ""
#For generira samo redak za ispis u trecoj tocki zadatka.
for i in range(len(n_total)):

    #If se odnosi na racun cetvrte tocke zadatka.
    if(n_total[i] >= znamenka_najvise_pojavljivanja):
        kolicina_najvise_pojavljivanja = n_total[i]
        znamenka_najvise_pojavljivanja = i
        
    redak_za_ispis += (str(i)+"-"+str(n_total[i]))
    if(i != len(n_total)-1):
        redak_za_ispis += ", "

print("\t\t"+redak_za_ispis)
print("Najviše se pojavljuje znamenka", znamenka_najvise_pojavljivanja, "i to",kolicina_najvise_pojavljivanja,"puta.")

#Peta tocka zadatka.
dat = open("zadatak-1.txt", "r")
for redak in dat:
    if str(znamenka_najvise_pojavljivanja) in redak:
        print(redak[:-1])
dat.close()


