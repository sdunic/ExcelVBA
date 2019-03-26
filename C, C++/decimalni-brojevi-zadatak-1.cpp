//Unesi realni broj od 1 do 100 tako da se maksimalno koristi 5 znamenki (npr. 12,345).
//Program treba ispisati umnožak sume znamenki cijelog dijela i decimalnog dijela ((1+2)*(3+4+5)).
//Na početku treba provjera je li broj u intervalu.

#include <iostream>
using namespace std;

bool provjeriRealniBroj(float realniBroj) {
    if(realniBroj < 1 || realniBroj >= 100) {
        return false;
    } else {

    }

    return true;
}

int zbrojiZnamenke(int n) {
    int rezultat = 0;

    while(n > 0) {
        rezultat += (n % 10);
        n = n / 10;
    }

    return rezultat;
}

int obradaBroja(float n) {
    int cijeliDio = int(n);
    int decimalniDio = (n - cijeliDio) * 10000;
    return zbrojiZnamenke(cijeliDio) * zbrojiZnamenke(decimalniDio);
}

int main() {
    float realniBroj = 12.345;

    cout << "Unesite realni broj od 1 do 100 tako da se maksimalno koristi 5 znamenki (npr. 12,345): ";
    cin >> realniBroj;

    if(provjeriRealniBroj(realniBroj)){
        cout << realniBroj << endl;
        cout << obradaBroja(realniBroj) << endl;
    } else {
        cout << "Neispravno unesen broj!";
    }

    return 0;
}
