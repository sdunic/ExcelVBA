/* 
	Napiši program koji ispisuje prvih 20 brojeva i na kraju ispište zbroj prvih 20 brojeva.
*/

#include <stdio.h>
void main()
{
	int x ;
	int zbroj=0;
	
	for(x = 1; x <= 20; x++)
	{
	 	printf("%d\n",x);
	 	zbroj +=x;
 	}
 	
	printf("Zbroj prvih 20 brojeva : %d",zbroj);
}
