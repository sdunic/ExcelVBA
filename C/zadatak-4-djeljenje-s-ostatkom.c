/*
	Napišite program u C u kojem je potrebno unijeti dva cijela broja te izraèunati cjelobrojni
	kvocijent i ostatak dijeljenja. Takoðer potrbno je ispitati jesu li ta dva broja djeljiva.
	Ispis neka bude oblika:
	Upisi djeljenik:
	Upisi djelitelj:
	... : ... = ...cijelih, a ostatak je ...
	Djeljivi = ... 
*/

#include<stdio.h>
void main()
{
	//deklaracija varijabli
	int a;
	int b;
	int cijelih;
	int ostatak;
	
	//unos i spremanje u varijable
	printf("Upisi djeljenik: ");
	scanf("%d", &a);
	printf("Upisi djelitelj: ");
	scanf("%d", &b);

	//racunske operacije
	cijelih = a / b;
	ostatak = a % b;
		
	//ispis na ekran
	printf("%d : %d = %d cijelih, a ostatak je %d\n", a,b,cijelih, ostatak);
	
	//dodatna provjera djeljivosti i ispis na ekran
	if(ostatak == 0){
		printf("Djeljivi = DA");
	}
	else{
		printf("Djeljivi = NE");
	}
}
