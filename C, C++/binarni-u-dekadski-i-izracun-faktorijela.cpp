//Binarni broj pretvoriti u dekadski i ispisati faktorijele od tog dekadskog.

#include <iostream>
#include <math.h>
using namespace std;

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

long faktorijel(int n){
    long faktorijel = 1;
    int i;

    for (i=1; i<=n; i++) {
        faktorijel *= i;
    }

    return faktorijel;
}

int main()
{
    long binarni = 111;

    int dekadski = binarniUDekadski(binarni);

    cout << dekadski << endl;
    cout << faktorijel(dekadski) << endl;

    return 0;
}

