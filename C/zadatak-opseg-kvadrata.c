/* 
	Zadatak: Uèitati stranicu kvadrata i spremiti u varijablu a (int). Nakon toga izraèunati opseg kvadrata u drugu varijablu o (float) po formuli OPSEGA 4*a. Za kraj ispisati opseg na ekran.
*/

#include <stdio.h>
void main()
{
	int a;
	float o;
	
	scanf("%d", &a);
	
	o=4*a;
	
	printf("Opseg kvadrata: %.4f", o);
	 
}
