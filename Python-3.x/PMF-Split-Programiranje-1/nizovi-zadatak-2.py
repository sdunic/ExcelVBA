#Upišite niz od n prirodnih brojeva.
#Ispisati samo one elemente niza kojima je suma znamenaka prost broj.

def prost(x):
    x = str(x)
    suma = 0
    
    for i in range(len(x)):
        suma += int(x[i])

    for i in range(2, suma):
        if suma%i==0:
            return False   
    return True

n = int(input("Unesite n: "))
niz = [0]*n

for i in range(n):
    niz[i] = int(input("Unesite element niza: "))

for i in range(n):
    if prost(niz[i]):
        print(niz[i])
