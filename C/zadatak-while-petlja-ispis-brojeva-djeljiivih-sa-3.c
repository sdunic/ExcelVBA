/*
	Uz pomoæ while petlje ispišite brojeve veæe od 100 i manje od 200 koji su djeljivi sa 3.
*/
#include <stdio.h>
void main()
{
	int x = 101;
	 
	while( x < 200)
	{
		if(x % 3==0)
		{ 
			printf("%d\n",x);
		}
		x++;
	}
}
