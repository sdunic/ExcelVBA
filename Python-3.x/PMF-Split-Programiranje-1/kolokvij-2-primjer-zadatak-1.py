#Upisati niz troznamenkastih prirodnih brojeva. Provjeriti ispravnost unosa. Ispisati broj pogrešno unesenih brojeva.
#Ispisati na zaslon samo one elemente niza kojima je produkt znamenki, umanjen za 1, djeljiv s 5, kao i produkte znamenki.
#Zadatak treba riješiti matematičkim metodama bez pretvaranja u stringove. Snimiti pod imenom zad1.py.

def ProvjeraIspravnosti(x):
    brojac = 0
    while x!=0:
        brojac += 1
        x=x//10
    if brojac == 3:
        return True
    return False

def ProvjeraIspisa(x):
    if (x-1)%5==0:
        return True
    return False

def ProduktZnamenki(x):
    produkt=1
    while x!=0:
        produkt *= x%10
        x= x//10
    return produkt


n = int(input("Unesi n: "))
niz = [0]*n

broj_pogresnih = 0

for i in range(n):
    niz[i] = int(input("Unesi broj u niz: "))

    if ProvjeraIspravnosti(niz[i]) == False:
        broj_pogresnih +=1

    #if niz[i] > 999 and niz[i] < 100:
        #broj_pogresnih +=1

print("Broj pogrešno upisanih:", broj_pogresnih)
print("Rješenje:")

for i in range(n):
    y = ProduktZnamenki(niz[i])
    
    #if (y-1)%5==0:
        #print(niz[i], "- produkt znamenki", y)
        
    if ProvjeraIspisa(y):
        print(niz[i], "- produkt znamenki", y)
            
    
