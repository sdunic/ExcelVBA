#Unesi tekst u datoteku. Ispiši samo velika slova koja se nalaze na prostom mjestu.

def Prosti(x):
    if x < 2:
        return False
    
    for i in range(2, x):
        if x%i==0:
            return False
        
    return True


dat = open("datoteka.txt", "w")
dat.write(input("Unesite rečenicu: "))
dat.close()

dat = open("datoteka.txt", "r")
recenica = dat.readline()

for i in range(len(recenica)):
    if recenica[i]>='A' and recenica[i]<='Z' and Prosti(i):
        print(recenica[i])

dat.close()
