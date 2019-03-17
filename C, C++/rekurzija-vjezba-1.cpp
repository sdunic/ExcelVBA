//50 puta ispisati "Dobar dan" bez petlji

#include <iostream>
using namespace std;

int ispis(int n)
{
    if(n==0)
        return 0;

    cout << "Dobar dan!" << endl;

    ispis(n-1);
    return 0;
}

int main()
{
    ispis(50);
    return 0;
}
