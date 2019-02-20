/*
Ispis zadane recenice s obrnutim poretkom rijeci.
*/
#include <stdio.h>
#include <string.h>

typedef struct cvor{
    char* tekst;
    struct cvor *iduci;
} rijec;

int main()
{
    rijec *pocetak = NULL;

	char str[] = "ja sam josip slipcevic";
	int init_size = strlen(str);
	char delim[] = " ";
	char *ptr = strtok(str, delim);

	while (ptr != NULL)
	{
	    rijec *nova = malloc(sizeof(rijec));

	    nova->tekst = ptr;
        nova->iduci = pocetak;
        pocetak = nova;

		ptr = strtok(NULL, delim);
	}

	while (pocetak != NULL)
	{
        printf("%s ", pocetak->tekst);
        pocetak = pocetak->iduci;
	}

	return 0;
}
