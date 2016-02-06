dat = open("zadatak-2.txt", "w")

ponovni_unos = ""
while ponovni_unos != "NE":
    rijec = input("Unesi riječ: ")
    dat.write(rijec + "\n")
    ponovni_unos = (input("Unosimo ponovno DA/NE - ")).upper()
dat.close()


dat = open("zadatak-2.txt", "r")
for redak in dat:
    br = 0
    for znak in redak:
        if znak >= 'A' and znak <= 'Z':
            br +=1
    if br%2 == 0 and br>0:
        print(redak, end="")
        
