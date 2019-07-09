#include <stdio.h>
#include <stdlib.h>
#include <string.h>

void main() {
     load_data();
}

void start() {
     printf("Uruchomiono program do autoprojektowania.\n");
     int a ;
     printf("dawaj liczbe ");
     scanf("%d\n", &a);
     printf("wpisałes:  %d\n", a);
     return 0;
}

void load_data()
{
     int T1;
     int one_person_data_number = 0;
     int number_of_persons = 0;
     char strs[13][20];
     char buffer[1024] ;
     char *record,*line;
     int i=0,j=0;
     int mat[100][100];
     FILE *fstream = fopen("\Book1.csv","r");
     if(fstream == NULL)
     {
          printf("\n file opening failed ");
          return -1 ;
     }
     while((line=fgets(buffer,sizeof(buffer),fstream))!=NULL)
     {
          record = strtok(line,";");
          while(record != NULL)
          {
               printf("\n record : %s",record) ;    //here you can put the record into the array as per your requirement.
               if(one_person_data_number==2){
                    T1 = record;
                    printf("\n%s\n", T1);
               }
               mat[i][j++] = atoi(record) ;
               record = strtok(NULL,";");
               ++one_person_data_number;
          }
          ++i ;
          ++number_of_persons;
     }
     printf("\n%s\n", T1);
     //printf("%d\n", one_person_data_number);
     //printf("%d\n", number_of_persons);
     return 0 ;
 }
