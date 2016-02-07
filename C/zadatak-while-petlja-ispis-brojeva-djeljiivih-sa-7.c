/*
	Uz pomoæ while petlje ispišite brojeve veæe od 500 i manje od 1700 koji su djeljivi sa 7.
*/

#include <stdio.h>
void main()
{
	int x = 501;
	
	while ( x < 1700 ) 
	{ 
		if(x % 7 == 0)
		{
    		printf( "%d\n", x );
		}
		x++;            
 	}
}

