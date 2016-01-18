dat = open("zadatak-3.txt", "w")

ponovni_unos=""
while ponovni_unos != "NE":
    recenica = input("Unesi rečenicu: ")
    if len(recenica) < 20:
        dat.write(recenica + "\n")
    ponovni_unos = (input("Unosimo ponovo DA/NE - ")).upper()
dat.close()
x = input("Unesi znak: ")


samoglasnici = ['a','e','i', 'o', 'u']


dat = open("zadatak-3.txt", "r")
tocka2 = ""
for redak in dat:
    for znak in redak:
        if znak != '\n':
            if znak.lower() in samoglasnici:
                tocka2 += (x + znak)
            else:
                tocka2 +=(znak)
    tocka2 +=(", ")

print(tocka2[:-2])
dat.close()



dat = open("zadatak-3.txt", "r")
for redak in dat:
    tocka3 = redak[:-1] + ":\t\t";
    n = [0, 0, 0, 0, 0, 0]
    for znak in redak:
        if znak != '\n':
            if znak.lower() in samoglasnici:
                n[samoglasnici.index(znak.lower())] += 1
            elif znak.lower() == x:
                n[-1] += 1
    for i in range(len(n)):
        if i == len(n)-1:
            tocka3 += ("znak "+ x + "-"+ str(n[-1]))
        else:
            tocka3 += (samoglasnici[i] + "-"+ str(n[i]) + ", ")
    print(tocka3)

dat.close()

#tocka4 ostaje i tocka5 zadatka
