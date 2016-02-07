/*
	Uz pomoæ for petlje ispišite brojeve veæe od 500 i manje od 1700 koji su djeljivi sa 7 i izraèunajte koliko ima takvih brojeva ukupno.
*/

#include <stdio.h>
void main()
{
	int x;
	int brojac = 0;
	
	for( x = 501; x < 1700; x++){
		if( x % 7 == 0 ){
			printf("%d\n", x);
			brojac++;
		}
	}
	
	printf("Ukupno brojeva ima: %d", brojac);

}

