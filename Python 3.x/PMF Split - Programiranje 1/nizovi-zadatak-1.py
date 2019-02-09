#Upišite niz od n prirodnih brojeva.
#Ispisati elemente kojima je suma znamenki paran broj.

def SumaZnamenki(x):
    suma = 0
    while x != 0:
        y = x%10
        suma += y
        x = x//10
    if suma % 2 == 0:
        return True
    return False

n=0
while n<=0:
    n=int(input("Unesite prirodan broj: "))
niz=[0]*n

for i in range(n):
    niz[i]=int(input("Unesite broj u niz: "))

for i in range(n):
    if SumaZnamenki(niz[i])==True:
        print(niz[i])
