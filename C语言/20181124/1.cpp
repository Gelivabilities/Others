#include <iostream>
using namespace std;
class math
{
private:
	const int num1=3;
	const int num2=7;
public:
	int add();
	int minus();
	int multiply();
	int divide();
};
math::add()
{
	return math.num1+math.num2;
}

int main()
{
	math a;
	a.add();
	return 0;
}

/*
class CMyString
{
public:
	CMyString(char *pData=NULL);
	CMyString(const CMyString &str);
	~CMyString(void);
private:
	char *m_pData;
};
CMyString& CMyString::operator=(const CMyString &str)
{
	if(this==&str)return *this;
	delete []m_pData;
	m_pData=NULL;
	m_pData=new char[strlen(str.m_pData)+1];
	strcpy(m_pData,str.m_pData);
}*/
