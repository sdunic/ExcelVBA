/*
	Napišite C program koji omoguæava korisniku da upiše dva realna broja. Treba izraèunati i ispisati kvadrat veæeg broja od ta dva. 
*/
#include <stdio.h>
void main()
{
	float x;
	float y;
	
	scanf("%f %f",&x,&y);
	
	if( x > y )
	{
	 	printf("Kvadrat veæeg broja %f je broj %f\n", x, x*x);
	}
	else
	{
		printf("Kvadrat veceg broja %f je broj %f\n",y,y*y);
	}
}
