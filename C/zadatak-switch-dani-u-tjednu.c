/* 

	Napišite program u C-u koji unosi cijeli broj u rasponu od 1 do 7 i primjenom SWITCH, za uneseni broj ispisuje odgovarajuæi dan u tjednu. Primjer:
	1 - ponedjeljak
	2 - utorak
*/


#include <stdio.h>
void main(){
	
	int x;
	
	scanf("%d",&x);
	
	switch(x) 
	{
      	case 1 :
        	printf("Ponedjeljak\n" );
        	break;
    	case 2 :
        	printf("Utorak\n" );
        	break;
    	case 3 :
        	printf("Srijeda\n" );
        	break;
        case 4 :
        	printf("Èetvrtak\n" );
        	break;
        case 5 :
        	printf("Petak\n" );
        	break;
        case 6 :
        	printf("Subota\n" );
        	break;
        case 7 :
        	printf("Nedjelja\n" );
        	break;
   }	
}
