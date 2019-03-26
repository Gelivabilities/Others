class GenderError(TypeError):
    pass

class animal():
    def __init__(self,name,gender,height,weight,age,type):
        if not isinstance(name,str) or not isinstance(gender,str) \
                or not isinstance(type,str) or not isinstance(height,int)\
                or not isinstance(age,int) or not isinstance(weight,int):
            raise TypeError
        if age<0 or height<=0 or weight<=0:
            raise ValueError
        if not (gender.lower() in ['公','母','男','女','雄','雌','male','female','man','woman']):
            raise GenderError
        self.name=name
        self.type=type
        self.gender=gender
        self.height=height
        self.weight=weight
        self.age=age
    def __str__(self):
        return 'Animal: '+self.type+', name: '+self.name+', gender: '+self.gender
    def eat(self,thing):
        print('I am a(n) '+self.type+', I am eating '+thing+'.')
    def hello(self):
        print('I am a(n) '+self.type+', my name is '+self.name+', aged '+str(self.age)+'.')

class human(animal):
    def __init__(self,name,gender,height,weight,age):
        animal.__init__(self,name,gender,height,weight,age,'human')
    def __str__(self):
        return 'Name: '+self.name
    def study(self,thing):
        if not isinstance(thing,str):
            raise TypeError
        print('I am '+self.name+', I am studying '+thing+'.')

class cat(animal):
    def __init__(self,name,gender,height,weight,age):
        animal.__init__(self,name,gender,height,weight,age,'cat')
    def __str__(self):
        return 'A cat named '+self.name+'.'
    def meow(self):
        print('Meow.')

class dog(animal):
    def __init__(self,name,gender,height,weight,age):
        animal.__init__(self,name,gender,height,weight,age,'dog')
    def __str__(self):
        return 'A dog named '+self.name+'.'
    def bark(self):
        print('汪！')

h=human('一个逗比','Female',175,69,24)
print(h)
h.eat('rice')
h.study('Python')
print('')
c=cat('大肥橘','公',36,24,2)
print(c)
c.eat('mouse')
c.meow()
print('')
d=dog('蛤士奇','雌',51,31,1)
print(d)
d.eat('shit')
d.bark()
print('')
print(isinstance(h,animal))
print(isinstance(c,human))
print(isinstance(d,dog))
#新型动物，还没有类，该怎么写？
a=animal('六耳猕猴','雄',180,70,550,'妖怪')
a.hello()