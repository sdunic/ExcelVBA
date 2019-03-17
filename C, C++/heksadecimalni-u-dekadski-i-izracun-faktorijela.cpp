//Binarni broj pretvoriti u dekadski i ispisati faktorijele od tog dekadskog.

#include <iostream>
#include <math.h>
#include <string.h>
using namespace std;


long heksadecimalniUBinarni(string heksadekadski)
{
    string hex = "0123456789ABCDEF";
    int bin [] = { 0, 1, 10, 11, 100, 101, 110, 111, 1000, 1001, 1010, 1011, 1100, 1101, 1110, 1111 };

    long binarni = 0;

    for (int i = 0; i < heksadekadski.size(); i++){
        for (int k = 0; k < hex.size(); k++){
            if( heksadekadski[i] == hex[k] ) {
                if(binarni == 0) {
                    binarni += bin[k];
                }
                else {
                    binarni = 10000*binarni + bin[k];
                }
            }
        }
    }

    return binarni;
}

int binarniUDekadski(long n)
{
    int dekadski = 0;
    int potencija = 0;

    while(n > 0){
        if ( n % 10 == 1 )
           dekadski += pow(2, potencija);

        potencija++;
        n = n / 10;
    }

    return dekadski;
}

long long int faktorijel(int n){
    long long int faktorijel = 1;
    int i;

    for (i=1; i<=n; i++) {
        faktorijel *= i;
    }

    return faktorijel;
}

int main()
{
    string heksadecimalni = "12";

    long binarni = heksadecimalniUBinarni(heksadecimalni);

    int dekadski = binarniUDekadski(binarni);

    cout << "HEX: " << heksadecimalni << endl;
    cout << "BIN: " << binarni << endl;
    cout << "DEC: " << dekadski << endl;
    cout << "FACT: " << faktorijel(dekadski) << endl;

    return 0;
}

