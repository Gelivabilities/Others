class A:
    def ping(self):
        print("ping in A")

class B(A):
    def pong(self):
        print("pong in B")

class C(A):
    def pong(self):
        print("PONG in C")

class D(B,C):#声明顺序不同，输出也不同
    def ping(self):
        super().ping()
        print("ping in D")

    def pingpong(self):
        self.ping()
        super().ping()
        self.pong()
        super().pong()
        C.pong(self)   # 在定义时调用特定父类的写法，显示传入self参数

d = D()
#最下面有答案
d.pong()
d.pingpong()











#答案1：B
#答案2：ADABBC
