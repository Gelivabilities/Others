import datetime

class PersonGenderError(TypeError):
    pass
class PersonValueError(TypeError):
    pass
class PersonNameError(TypeError):
    pass
class PersonTypeError(TypeError):
    pass
class CourseError(TypeError):
    pass

class Person:
    _num=0

    def __init__(self,name,sex,birthday,ident):
        if not isinstance(name, str)and sex in ['女','男']:
            raise PersonGenderError
        try:
            birth=datetime.date(*birthday)
        except:
            raise PersonValueError
        self._name=name
        self._sex=sex
        self._birthday=birth
        self._id=ident
        Person._num+=1

    def id(self):return self._id
    def name(self):return self._name
    def sex(self):return self._sex
    def birthday(self):return self._birthday
    def age(self):
        #判断年龄不只是要年份，还得看月日
        judge1=datetime.date.today().month>self._birthday.month
        judge2=datetime.date.today().day>=self._birthday.day
        judge3=datetime.date.today().month==self._birthday.month
        if judge1 or (judge3 and judge2):
            return(datetime.date.today().year-self._birthday.year)
        return (datetime.date.today().year-self._birthday.year)-1
    def set_name(self,input):
        if not isinstance(input,str):raise PersonNameError
        self._name=input
    def __lt__(self,other):
        if not isinstance(other,Person):raise PersonTypeError
        return self._id<other._id

    @classmethod
    def num(cls):return Person._num

    def __str__(self):
        return ' '.join((str(self._id),self._name,self._sex,str(self._birthday)))

    def details(self):
        return '，'.join(('编号：'+str(self._id),'姓名：'+self._name,
                         '性别：'+self._sex,'年龄：'+str(self.age()),'出生日期：'+str(self._birthday)))
'''
p1 = Person('张三', '男', (1995, 4, 20), '015')
p2 = Person('李四', '女', (1994, 3, 31), '079')
p3 = Person('王五', '女', (1998, 11, 5), '003')
p4 = Person('赵六', '男', (1993, 7, 26), '044')

plist=[p1,p2,p3,p4]
for p in plist:
    print(p)
print('')

plist.sort()
for p in plist:
    print(p.details())
'''

class Student(Person):
    _id_num=0

    @classmethod
    def _id_gen(cls):
        cls._id_num+=1
        return cls._id_num

    def __init__(self,name,sex,birthday,department):
        Person.__init__(self,name,sex,birthday,Student._id_gen())
        self._department=department
        self._enroll_date=datetime.date.today()
        self._courses={}
    def set_course(self,course_name):
        self._courses[course_name]=None

    def set_score(self,course_name,score):
        if course_name not in self._courses:
            raise CourseError
        self._courses[course_name]=score
    def scores(self):return [(c+': '+self._courses[c])for c in self._courses]
    def details(self):
        return '，'.join((Person.details(self),'入学日期：'+str(self._enroll_date),
                         '院系'+self._department,'\n成绩：'+str(self.scores())))

s1=Student('陈七','女',(1995,4,20),'计算机系')
print(s1)
print(s1.details())