/*
	Program koji ucitava dva broja s tipkovnice. Zatim ih zbraja i na kraju ispusuje na ekran.
*/
#include <stdio.h>

void main()
{
	//deklaracija varijabli
	int a, b, c;

	//unos i spremanje u varijable
	scanf("%d %d", &a, &b);
	
	//racunske operacije
	c = a + b;
	
	//ispis zbroja na ekran
	printf("Zbroj brojeva je %d", c);
}
