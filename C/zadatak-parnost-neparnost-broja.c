/*
	1. Napisite program koji ispituje (pomoæu modul operatora %) da li je upisani broj paran ili neparan.
	2. Prosirite zadatak 1 za provjeru parnosti, tako da za parne i neparne brojeve dodatno ispitate 
	djeljivost s 3, a  za parne brojeve i djeljivost s 4.
*/

#include <stdio.h>

void main()
{
	//deklaracija varijabli
	int a;
	
	//unos i spremanje u varijable
	scanf("%d", &a);
	
	//logièke provjere parnosti s proširenjem zadatka te ispis na ekran
	if(a % 2 == 0)
	{
		printf("Broj je paran!\n");
		if(a % 3 == 0){
			printf("Broj je djeljiv s 3!\n");
		}
		if(a % 4 == 0){
			printf("Broj je djeljiv s 4!\n");
		}
		
	}
	else
	{
		printf("Broj je neparan!\n");
		if(a % 3 == 0){
			printf("Broj je djeljiv s 3!\n");
		}
	}
}
