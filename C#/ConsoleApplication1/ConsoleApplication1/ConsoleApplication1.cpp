// ConsoleApplication1.cpp : 定义控制台应用程序的入口点。
//

#include "stdio.h"
#include "stdlib.h"


int main()  
{
   int a,b,c,max;  
   scanf("%d,%d,%d",&a,&b,&c); /*请输入三个数*/ 
   max=a;   
   if (a<b){max=b;}
   if (b<c){max=c;}
   printf("max(%d,%d,%d)=%d\n",a,b,c,max);
   system("pause");
   return 0;
}

