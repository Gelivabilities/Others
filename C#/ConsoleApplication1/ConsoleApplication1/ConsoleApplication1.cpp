// ConsoleApplication1.cpp : �������̨Ӧ�ó������ڵ㡣
//

#include "stdio.h"
#include "stdlib.h"


int main()  
{
   int a,b,c,max;  
   scanf("%d,%d,%d",&a,&b,&c); /*������������*/ 
   max=a;   
   if (a<b){max=b;}
   if (b<c){max=c;}
   printf("max(%d,%d,%d)=%d\n",a,b,c,max);
   system("pause");
   return 0;
}

