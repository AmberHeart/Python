import xlwings as xw
import random


class classify:

    def work(choice, groupnum):

        class Student():

            def __init__(self, num, name, grade, college, classs):
                self.num = num
                self.name = str(name)
                self.gender = 0  # 默认性别0为男性
                self.grade = grade
                self.college = college
                self.classs = classs
                if self.name[len(str(self.name)) - 1] == '*':
                    self.gender = 1  # 性别1 为女性
                    # self.name = self.name - '*'

        # 学生数据顺序分别是学号 名字 年级 学院 班级
        stunum = 0  # 学生总数
        stulist = []  # 存储每个学生数据
        # 当前App下新建一个Book
        # visible参数控制创建文件时可见的属性
        app = xw.App(visible=False, add_book=False)
        input = app.books.open("答疑分组.xlsx")  # 原始数据
        output = app.books.add()  # 生成excel
        # 实例化一个工作表对象
        isheet1 = input.sheets["sheet0"]
        osheet1 = output.sheets["sheet1"]
        # 读入学生数据
        n = 65
        x = 2
        n = chr(n)  # ASCII字符
        while isheet1.range('A%d' % x).value != None:
            stulist.append(
                Student(
                    isheet1.range('B%d' % x).value,
                    isheet1.range('C%d' % x).value,
                    isheet1.range('D%d' % x).value,
                    isheet1.range('E%d' % x).value,
                    isheet1.range('F%d' % x).value))
            stunum += 1  # 记录学生总数
            x += 1
        # 写入分组数据

        name = []
        number = []
        grade = []
        faculty = []
        gender = []
        randnum = []
        n = stunum #总人数
        cnt = int(groupnum) #每组人数
        kind = choice #分类方法
        kindnum = 0 #种类数
        kindpos = [] #每个种类的第一个学生所在位置
        kindcnt = [] #每个种类有多少人
        now_kindpos = [] #现在各个种类分配数字到第几个人

        # 读入name number kind faculty gender randnum cnt 跟Student类对应
        for i in range(0, n):
            name.append(stulist[i].name)
            number.append(stulist[i].num)
            grade.append(stulist[i].grade)
            gender.append(stulist[i].gender)
            faculty.append(stulist[i].college)
            randnum.append(random.randint(0, n))

        #交换函数
        def swapstudent(x, y):
            tmp = name[x]
            name[x] = name[y]
            name[y] = tmp
            tmp = number[x]
            number[x] = number[y]
            number[y] = tmp
            tmp = faculty[x]
            faculty[x] = faculty[y]
            faculty[y] = tmp
            tmp = grade[x]
            grade[x] = grade[y]
            grade[y] = tmp
            tmp = gender[x]
            gender[x] = gender[y]
            gender[y] = tmp
            tmp = randnum[x]
            randnum[x] = randnum[y]
            randnum[y] = tmp
        #按完全随机的数字进行排列 打乱原有顺序
        for i in range(0, n):
            for j in range(i, n):
                if randnum[i] > randnum[j]:
                    swapstudent(i, j)
        #如果不是完全随机 对种类进行排序 让种类相同的排在一起
        if kind != 0:
            for i in range(0, n):
                for j in range(i, n):
                    if kind == 1:
                        if grade[i] > grade[j]:
                            swapstudent(i, j)
                    if kind == 2:
                        if faculty[i] > faculty[j]:
                            swapstudent(i, j)
                    if kind == 3:
                        if gender[i] > gender[j]:
                            swapstudent(i, j)
                            
        if kind != 0:
            kindpos.append(0)
            #按种类分组 求每个种类的参数
            if kind == 1:
                now_kind = grade[0]
                for i in range(0, n):
                    if now_kind != grade[i]:
                        kindpos.append(i)
                        kindnum += 1
                        now_kind = grade[i]
            elif kind == 2:
                now_kind = faculty[0]
                for i in range(0, n):
                    if now_kind != faculty[i]:
                        kindpos.append(i)
                        kindnum += 1
                        now_kind = faculty[i]
            elif kind == 3:
                now_kind = gender[0]
                for i in range(0, n):
                    if now_kind != gender[i]:
                        kindpos.append(i)
                        kindnum += 1
                        now_kind = gender[i]
            kindpos.append(n)
            kindnum += 1
            for i in range(0, kindnum):
                kindcnt.append((kindpos[i + 1] - kindpos[i]) * cnt // n + 1)#求按比例分每个组应有多少某个种类的人
                now_kindpos.append(kindpos[i])
            #分配排列序号
            nowi = 0
            while True:
                flag = 0
                for i in range(0, kindnum):
                    if kindpos[i + 1] < now_kindpos[i] + kindcnt[i]:
                        flag = 1
                if flag == 1:
                    break
                nowj = 0
                for i in range(0, kindnum):
                    for j in range(0, kindcnt[i]):
                        randnum[now_kindpos[i]] = nowi * cnt + nowj
                        now_kindpos[i] += 1
                        nowj += 1
                nowi += 1
            #按分配好的序号排序 未分配序号的学生以随机序号插入队中
            for i in range(0, n):
                for j in range(i, n):
                    if randnum[i] > randnum[j]:
                        swapstudent(i, j)

        rest = n % cnt #多出的学生数
        cnt += 1
        i = 0
        x = 1
        #分组
        for zushu in range(0, n):
            if rest == 0:
                rest = -1
                cnt -= 1
            if i >= n:
                break
            osheet1.range('A%d' % x).value = "组" + str(zushu) + "."
            x += 1
            for j in range(0, cnt):
                osheet1.range('A%d' % x).value = number[i + j]
                osheet1.range('B%d' % x).value = name[i + j]
                x += 1
            i += cnt
            if rest > 0:
                rest -= 1
            

        # 读值并打印
        # print('value of A1:', isheet1.range('C3').value)

        output.save('分组结果.xlsx')
        input.close()
        output.close()
        # 结束进程
        app.quit()
