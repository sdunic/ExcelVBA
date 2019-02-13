#tekst1.txt je tekstualna datoteka u kojoj su sadržane riječi odvojene razmakom.
#Napišite program koji će stvoriti novu datoteku tekst2.txt u koju ćete u prvom redu ispisati broj riječi iz polazne datoteke,
#a u svim redovima u nastavku ispisati duljinu svake riječi iz prve datoteke.

dat1 = open("tekst1.txt","r")
dat2 = open("tekst2.txt","w")

rec = dat1.readline()

brojac=1
for znak in rec:
    if znak == " ":
         brojac+=1

dat2.write("Broj riječi: " + str(brojac) + "\n")

brojac=0
for znak in rec:
    brojac += 1
    if znak == " ":
        brojac -= 1
        dat2.write(str(brojac)+"\n")
        brojac = 0

dat1.close()
dat2.close()
