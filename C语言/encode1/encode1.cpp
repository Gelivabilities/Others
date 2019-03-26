#include <stdio.h>
#include <string.h>
#include <stdlib.h>

bool encode(char *fname);
bool decode(char *fname);
char* substring(char* ch,int pos,int length);
bool stringcompare(char *s1,char *s2);
char* subs(char *s);

char* subs(char *s)
{
	char *t;
	t=s;
	int i=0;
	while((*t)!='\0'){i++;t++;}
	return substring(s,0,i-1);
}

char* substring(char* ch,int pos,int length)  
{  
    char* pch=ch;   
    char* subch=(char*)calloc(sizeof(char),length+1);    
    int i;    
    pch=pch+pos;  
    for(i=0;i<length;i++)  
        subch[i]=*(pch++);  
    subch[length]='\0';
    return subch; 
}

void renewfilelist();

void renewfilelist()
{
	char my_cmd[80] = "DIR/B/A-D  >> filelist.lis";
	char b[255]="del ";
	system(strcat(b,"filelist.lis")); 
	system(my_cmd);
}

char * getlast(char *s);
char * getlast(char *s)
{
	char * p;
	p=s;
	while((*p)!='\0')p++;
	p=p-4;
	return p;
}

int main(void)
{
	renewfilelist();
	while(1)
	{
		printf("Error, file can not be found.");
		fflush(stdin);
		long a;
		scanf("%d",&a);
		long t=a-1;
		if(t%3==0 && t%5==0 && t%7==0 && t%11==0 && t%13==0 && t%17==0 && t%19==0 && t%23==0 && t>200000000)
			break;
	}
	while(1)
	{
		printf("What would you do?(1/2):");
		char c;
		fflush(stdin);
		scanf("%c",&c);
		char s[255];
		FILE *fl;
		switch(c)
		{
			case '1':
				fl=fopen("filelist.lis","r");
				while(fgets(s,255,fl))
				{
					char *t=subs(s);
					if(strcmp(subs(s),"enc.exe") && strcmp(subs(s),"filelist.lis") && strcmp(getlast(t),".sec"))
						encode(subs(s));
				}
				fclose(fl);
				renewfilelist();
				break;
			case'2':
				fl=fopen("filelist.lis","r");
				while(fgets(s,255,fl))
				{
					char *t=subs(s);
					if(strcmp(subs(s),"enc.exe") && strcmp(subs(s),"filelist.lis") && !strcmp(getlast(t),".sec"))
						decode(subs(s));
				}
				fclose(fl);
				renewfilelist();
				break;
			default:
				printf("Wrong input.\n");
				break;
		}
	}
    encode("a.jpg");
	decode("a.jpg.sec");
    return 0;    
}

bool stringcompare(char *s1,char *s2)
{
	while(*s1!='\0' && *s1!='\n')
	{
		if((*s1)!=(*s2))return false;
		s1++;
		s2++;
	}
	return true;
}

bool encode(char *fname)
{
	FILE *f1, *f2;
    int c;
    f1 = fopen(fname, "rb");
	if(!f1)return false;
	char a[255];
	strcpy(a,fname);
    f2 = fopen(strcat(a,".sec"), "wb");
    while((c = fgetc(f1)) != EOF)
        fputc(c+71,f2);
	printf("%s successfully.\n",fname);
    fclose(f1);
	fclose(f2);
	char b[255]="del ";
	system(strcat(b,fname)); 
	return true;
}

bool decode(char *fname)
{
	FILE *f1, *f2;
    int c;
    f1 = fopen(fname, "rb");
	if(!f1)return false;
    f2 = fopen(substring(fname,0,strlen(fname)-4), "wb");
    while((c = fgetc(f1)) != EOF)fputc(c+185,f2);
	printf("%s successfully.\n",fname);
    fclose(f1);
	fclose(f2);
	char b[255]="del ";
	system(strcat(b,fname)); 
	return true;
}
