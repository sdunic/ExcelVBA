/*
	Uz pomoæ for petlje ispišite brojeve veæe od 100 i manje od 200 koji su djeljivi sa 3 i izraèunajte koliko ima takvih brojeva ukupno.
*/

#include <stdio.h>
void main()
{
	int x;
	int brojac=0;
	
	for(x = 101; x < 200;  x++ )
	{
		if(x % 3 == 0)
		{
			printf("%d\n",x);
			brojac++;
		}			
	}
	printf("ukupno brojeva ima: %d", brojac);
}
	


