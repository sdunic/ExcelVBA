/* 

	Napišite program u C-u koji unosi cijeli broj u rasponu od 1 do 5 i primjenom SWITCH, za uneseni broj ispisuje odgovarajuæi dan u tjednu. Primjer:
	1 - Nedovoljan
	2 - Dovoljan
*/

#include <stdio.h>
void main()
{
	int x; 
	
	scanf("%d",&x);
	
	switch (x)
	{ 
	    case 1:
			printf("nedovoljan\n");
			break ;
		case 2:
			printf("dovoljan\n");
			break ;
		case 3:
			printf("dobar\n");
			break ;
		case 4:
			printf("vrlo dobar\n");
			break ;
		case 5:
			printf("odlican\n");
			break ;
	}
}
