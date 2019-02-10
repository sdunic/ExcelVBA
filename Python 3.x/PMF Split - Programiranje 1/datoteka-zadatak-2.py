#Unesi u datoteku proizvoljan broj riječi.
#Ispisati iz datoteke riječi koje imaju više od 2 samoglasnika



dat = open("datoteka.txt", "w")

while True:
    dat.write(input("Unesi rijec u datoteku: ") + "\n")

    odg = input("Za prestanak pretisnite *, za nastavak pritisnite enter: ")
    if odg == "*":
        break

dat.close()

dat = open("datoteka.txt", "r")

for rijec in dat:
    brojac = 0
    for slovo in rijec[:-1]:
        if slovo.upper() in ['A', 'E', 'I', 'O', 'U']:
            brojac += 1
    if brojac >= 2:
        print(rijec[:-1])
        
dat.close()
    
