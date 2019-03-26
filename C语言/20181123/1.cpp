#include <stdio.h>
#include <iostream>
using namespace std;

/*
bool xor(bool a,bool b);
bool find(int **matrix,int rows,int cols,int num);

bool xor(bool a,bool b){
	return a&&!b || !a&&b;
}

bool find(int **matrix,int rows,int cols,int num)
{
	if (!rows*cols)return false;//行或列数为0
	int x=rows-1,y=0;//左下角起步
	if(*((int *)matrix+x*rows+y)==num) return true;//如果左下角元素就是它，直接返回true
	//先去行还是先去列，取决于左下角大于它还是小于它
	bool flag=*((int *)matrix+x*rows+y)>num;
	//去掉小于它的列和大于它的行
	while(xor(flag,*((int *)matrix+x*rows+y)<num))
	{
		if(flag?x==0:y>=cols-1)break;
		flag?x--:y++;
		if(*((int *)matrix+x*rows+y)==num) return true;
	}
	while(xor(flag,*((int *)matrix+x*rows+y)>num))
	{
		if(flag?y>=cols-1:x==0)break;
		flag?y++:x--;
		if(*((int *)matrix+x*rows+y)==num) return true;
	}
	return *((int *)matrix+x*rows+y)==num?true:false;
}

int main(){
	int matrix[4][4]={{1,2,8,9},{2,4,9,12},{4,7,10,13},{6,8,11,15}};
	int **p;
	p=(int **)matrix;
	for(int i=1;i<=16;i++)
		cout<<find((int **)matrix,4,4,i)<<endl;
}*/

/*
class CMyString
{
public:
	CMyString(char *pData=NULL);
	CMyString(const CMyString& str);
	~CMyString(void);
private: 
	char * m_pData;
};

CMyString& CMyString::operator=(const CMyString &str)
{
	if(this==&str)return *this;
	delete[]m_pData;
	m_pData=new char[strlen(str.m_pData)+1];
	strcpy(m_pData,str.m_pData);
	
	return *this;
}*/

/*
class C{
private:
	int int_value;
	float float_value;
	char char_value;
	char *string_value;
public:
	C(int n){int_value=n;}
	C(char c){char_value=c;}
	C(char *s){string_value=s;}
	C(float f){float_value=f;}
	void print_int(){cout<<int_value<<endl;}
	void print_float(){cout<<float_value<<endl;}
	void print_char(){cout<<char_value<<endl;}
	void print_string(){cout<<string_value<<endl;}
	void print_all(){
		print_int();
		print_float();
		print_char();
		print_string();
	}
};

int main(){
	C c=10;
	c.print_int();
	c='a';
	c.print_char();
	c="test";
	c.print_string();
	c=(float)2.333;
	c.print_float();
	return 0;
}*/



class threat//线程
{
private:
	int num;//线程编号
public:
	threat(int n){num=n;}
	threat(threat &old_threat){num=old_threat.num+1;}//拷贝的线程编号要加一
	Print(){cout<<num<<endl;}
};

int main(){
	threat a=10;
	//b和c复制a
	threat b=a;
	threat c=a;
	a.Print();
	b.Print();
	c.Print();
	//d复制b
	threat d=b;
	d.Print();
	return 0;
}


/*
class A
{
private:
	int value;

public:
	A(int n){cout<<"普通赋值:";value=n;}
	A(const A &other){cout<<"类型被复制，复制者要加一:";value=other.value+1;}

	void Print(){cout<<value<<endl;}
};

int main(){
	A a=10;
	a.Print();
	A b=a;
	b.Print();
	A c=b;
	c.Print();
	return 0;
}*/
