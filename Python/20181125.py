class Counterable():
    counter=0
    def __init__(self):
        Counterable.counter+=1
    def __str__(self):
        return str(Counterable.counter)
    def __int__(self):
        return Counterable.counter
    @staticmethod
    def get_count():
        print( '这是第'+str(Counterable.counter)+'个这种类型的变量')
'''
for i in range(10):
    a=Counterable()
    print(a)
    a.get_count()
'''
import numpy as np
class rational:
    @staticmethod
    def _gcd(m,n):
        if n==0:
            m,n=n,m
        while m!=0:
            m,n=n%m,m
        return n

    def __init__(self,num,den=1):
        if not isinstance(num,int) or not isinstance(den,int):
            raise TypeError
        if den==0:
            raise ZeroDivisionError
        self.__num=int(abs(num/self._gcd(num,den)))
        self.__den=int(abs(den/self._gcd(num,den)))
        self.value=num/den
        self.__sign=2*((num>=0 and den>=0) or (num<0 and den<0)+0)-1
        if self.__sign==1:
            sign_string=''
        else:
            sign_string='-'
        if self.__num==0 or self.__den==1:
            rear_str=''
        else:
            rear_str='/' + str(self.__den)
        self.value_string = sign_string+str(self.__num) + rear_str

    def __add__(self,r):
        new__num=self.__num*r.__den*self.__sign+r.__num*self.__den*r.__sign
        new__den=self.__den*r.__den
        return rational(new__num,new__den)
    def __sub__(self,r):
        new__num=self.__num*r.__den*self.__sign-r.__num*self.__den*r.__sign
        new__den=self.__den*r.__den
        self.__init__(new__num,new__den)
        return rational(new__num, new__den)
    def __mul__(self,r):
        new__num=self.__num*r.__num*self.__sign*r.__sign
        new__den=self.__den*r.__den
        return rational(new__num, new__den)
    def __truediv__(self,r):
        new__num=self.__num*r.__den*self.__sign*r.__sign
        new__den=self.__den*r.__num
        self.__init__(new__num, new__den)
        return rational(new__num, new__den)
    def __eq__(self,r):
        return self.__num*r.__den*self.__sign==r.__num*self.__den*r.__sign
    def __ne__(self,r):
        return self.__num*r.__den*self.__sign!=r.__num*self.__den*r.__sign
    def __lt__(self,r):
        if self.__sign > r.__sign: return False
        if self.__sign < r.__sign: return True
        p=self.__num*r.__den
        q=r.__num*self.__den
        return (p < q) ==(self.__sign==1) and p!=q
    def __gt__(self,r):
        if self.__sign > r.__sign: return True
        if self.__sign < r.__sign: return False
        p=self.__num*r.__den
        q=r.__num*self.__den
        return (p > q) ==(self.__sign==1) and p!=q
    def __le__(self,r):
        if self.__sign > r.__sign: return False
        if self.__sign < r.__sign: return True
        p=self.__num*r.__den
        q=r.__num*self.__den
        return (p < q) ==(self.__sign==1)
    def __ge__(self,r):
        if self.__sign > r.__sign: return True
        if self.__sign < r.__sign: return False
        p=self.__num*r.__den
        q=r.__num*self.__den
        return (p > q) ==(self.__sign==1)
    def __str__(self):
        return self.value_string
'''
r=rational(2,3)
print('有理数：'+r.value_string+'，小数值：'+str(r.value))
r=r+rational(6,7)
print('有理数：'+r.value_string+'，小数值：'+str(r.value))
r=r-(rational(14,-21))
print('有理数：'+r.value_string+'，小数值：'+str(r.value))
r=r*(rational(-3,1))
print('有理数：'+r.value_string+'，小数值：'+str(r.value))
r=r/rational(3,5)
print('有理数：'+r.value_string+'，小数值：'+str(r.value))
r=r*(rational(-21,1))
print('有理数：'+r.value_string+'，小数值：'+str(r.value))
r=r/rational(-13,4)
print('有理数：'+r.value_string+'，小数值：'+str(r.value))
r=rational(-13,4)
print(r>rational(14,6))
print(r==rational(-13,5))
print(r>rational(14,-4))

#字符串测试
a=rational(961,3)
b=rational(456)
print(b-a)
print(rational(1,2)+rational(3,4)*rational(5,6))
'''

class test():
    def __init__(self):
        x = Counterable()
        y = Counterable()
        self.out=rational(x.counter,y.counter)
    def __str__(self):
        return str(self.out)
    @staticmethod
    def get_count_rational(c1,c2):
        return rational(int(c1),int(c2))

for i in range(10):
    print(test.get_count_rational(i,i+1),end=' ')
print('')
for i in range(10):
    print(Counterable(),end=' ')
print('')

for i in range(10):
    print(test(),end=' ')

'''
for i in range(9):
    for j in range(i):
        print('|', end='')
        print(rational(j+1,i).value_string+'| ',end='')
    print('\n')
'''

#import time
'''
def test1(n):#极慢
    lst=[]
    for i in range(n):
        lst=lst+[i]#为什么慢，因为每次都要将很大个的lst复制出来然后赋值，所以复杂度变成了O(n2)，很垃圾
    return lst

def test2(n):#快很多
    lst=[]
    for i in range(n):
        lst.append(i)#这个append肯定不是复制出来再赋值的，所以还是线性复杂度，下面两个就不知道为什么这么快了
    return lst

def test3(n):#更快
    return [i for i in range(n)]

def test4(n):#最快
    return list(range(n))

n=30000;
print('开始')
time_start=time.time()
test1(n)
print('用时'+str(time.time()-time_start)+'s')
time_start=time.time()
test2(n)
print('用时'+str(time.time()-time_start)+'s')
time_start=time.time()
test3(n)
print('用时'+str(time.time()-time_start)+'s')
time_start=time.time()
test4(n)
print('用时'+str(time.time()-time_start)+'s')
'''