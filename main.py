#coding:utf-8
import tkinter as tk
import tkinter.messagebox 
import xlwt
import xlrd
import numpy as np
import matplotlib.pyplot as plt
from tkinter import filedialog

page = 0
sheetFrame = False
#config
collegeCol = 0 #学院名
majorCol = 1 #专业名
classCol = 2 #班级名
studentNameCol = 3
studentIdCol = 4
timeCol = 5 #学期
lessonIdCol = 6
lessonNameCol = 7
creditCol = 8 #学分
gradeCol = 9
titleLine = 1
pageSize = 30 #每页显示的信息条数
#jxnu 标准成绩单

#功能
def getLesson():
    lesson = {}
    for row in stu:
        s = row[lessonIdCol]+' '+row[lessonNameCol]
        if s not in lesson.keys():
            lesson[s] = [row]
        else:
            lesson[s].append(row)
    return lesson

def getStudent():
    student = {}
    for row in stu:
        s = row[studentIdCol]+' '+row[studentNameCol]
        if s not in student.keys():
            student[s] = [row]
        else:
            student[s].append(row)
    return student

def avg(list):
    return sum(list)/len(list)

def weightavg(list):
    ans = []
    for s in list:
        weight = 0.0 # 总学分
        sum = 0.0 # 学分乘成绩的和
        for x in s:
            weight += x[0]
            sum += x[0]*x[1]
        ans.append(sum/weight)
    return ans

def browse_file():
    tkFilename.set(filedialog.askopenfilename(filetypes=[('Excel file','.xlsx')]))

def exit_program():
    if tk.messagebox.askokcancel('Exit','Do you want to exit?'):
        root.destroy()

#关于
def show_about():
    top = tk.Toplevel()
    top.title('About')
    top.geometry('200x100')
    tk.Label(top,text='All Right Reserved').pack()
    tk.Label(top,text='Version 1.2.0').pack()

#设置
def set_variable():
    def submit_variable():
        global collegeCol,majorCol,classCol,studentNameCol,studentIdCol,timeCol,lessonIdCol,lessonNameCol,creditCol,gradeCol,pageSize
        collegeCol = tkcollegeCol.get()
        majorCol = tkmajorCol.get()
        classCol = tkclassCol.get()
        studentNameCol = tkstudentNameCol.get()
        studentIdCol = tkstudentIdCol.get()
        timeCol = tktimeCol.get()
        lessonIdCol = tklessonIdCol.get()
        lessonNameCol = tklessonNameCol.get()
        creditCol = tkcreditCol.get()
        gradeCol = tkgradeCol.get()
        pageSize = tkpageSize.get()
        tk.messagebox.showinfo(title='Set System Variable',message='修改成功，将在打开新表后生效！')
        top.destroy()

    tkcollegeCol = tk.IntVar()
    tkcollegeCol.set(collegeCol)
    tkmajorCol = tk.IntVar()
    tkmajorCol.set(majorCol)
    tkclassCol = tk.IntVar()
    tkclassCol.set(classCol)
    tkstudentNameCol =tk.IntVar()
    tkstudentNameCol.set(studentNameCol)
    tkstudentIdCol = tk.IntVar()
    tkstudentIdCol.set(studentIdCol)
    tktimeCol = tk.IntVar()
    tktimeCol.set(timeCol)
    tklessonIdCol = tk.IntVar()
    tklessonIdCol.set(lessonIdCol)
    tklessonNameCol = tk.IntVar()
    tklessonNameCol.set(lessonNameCol)
    tkcreditCol = tk.IntVar()
    tkcreditCol.set(creditCol)
    tkgradeCol = tk.IntVar()
    tkgradeCol.set(gradeCol)
    tkpageSize = tk.IntVar()
    tkpageSize.set(pageSize)
    top = tk.Toplevel()
    top.title('Set System Variable')
    tk.Label(top,text='文件信息(若没有相应信息则不修改)').grid(row = 0)
    tk.Label(top,text='学院名所在列').grid(row = 1)
    tk.Entry(top,textvariable = tkcollegeCol).grid(row=1,column=1)
    tk.Label(top,text='专业名所在列').grid(row = 2)
    tk.Entry(top,textvariable = tkmajorCol).grid(row=2,column=1)
    tk.Label(top,text='班级所在列').grid(row = 3)
    tk.Entry(top,textvariable = tkclassCol).grid(row=3,column=1)
    tk.Label(top,text='学生姓名所在列').grid(row = 4)
    tk.Entry(top,textvariable = tkstudentNameCol).grid(row=4,column=1)
    tk.Label(top,text='学生学号所在列').grid(row = 5)
    tk.Entry(top,textvariable = tkstudentIdCol).grid(row=5,column=1)
    tk.Label(top,text='学期所在列').grid(row = 6)
    tk.Entry(top,textvariable = tktimeCol).grid(row=6,column=1)
    tk.Label(top,text='课程号所在列').grid(row = 7)
    tk.Entry(top,textvariable = tklessonIdCol).grid(row=7,column=1)
    tk.Label(top,text='课程名所在列').grid(row = 8)
    tk.Entry(top,textvariable = tklessonNameCol).grid(row=8,column=1)
    tk.Label(top,text='学分所在列').grid(row = 9)
    tk.Entry(top,textvariable = tkcreditCol).grid(row=9,column=1)
    tk.Label(top,text='成绩所在列').grid(row = 10)
    tk.Entry(top,textvariable = tkgradeCol).grid(row=10,column=1)
    tk.Label(top,text='显示信息').grid(row = 11)
    tk.Label(top,text='每页显示信息条数').grid(row = 11)
    tk.Entry(top,textvariable = tkpageSize).grid(row=11,column=1)
    tk.Button(top,text='确定',command=submit_variable).grid(row=12)

#展示
def show_view(list):
    global page
    page = 0
    viewWindow = tk.Toplevel()
    viewWindow.title('Lesson')
    tk.Button(viewWindow,text='OK!',command=lambda:show_sheet(list[viewListbox.get(viewListbox.curselection())],sheettitle)).pack()
    listFrame = tk.Frame(viewWindow)
    scrollbar = tk.Scrollbar(listFrame, orient=tk.VERTICAL)
    viewListbox = tk.Listbox(listFrame, yscrollcommand=scrollbar.set)
    for s in list.keys():
        viewListbox.insert('end',s)
    scrollbar.config(command=viewListbox.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    viewListbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
    listFrame.pack()

def show_lesson():
    lesson = getLesson()
    show_view(lesson)

def show_student():
    student = getStudent()
    show_view(student)

def getLessonName(list):
    lessonName = []
    for x in list:
        s = x[lessonIdCol]+' '+x[lessonNameCol]
        if s not in lessonName:
            lessonName.append(s)
    return lessonName

#画图
def autolabel(rects,ax,xpos='center'):
    xpos = xpos.lower()  # normalize the case of the parameter
    ha = {'center': 'center', 'right': 'left', 'left': 'right'}
    offset = {'center': 0.5, 'right': 0.57, 'left': 0.43}  # x_txt = x + w*off
    for rect in rects:
        height = rect.get_height()
        ax.text(rect.get_x() + rect.get_width()*offset[xpos], 1.01*height,'{}'.format(height), ha=ha[xpos], va='bottom')

def draw_barchart(list,chartname):
    ind = np.arange(len(list))  # the x locations for the groups
    width = 0.7
    fig, ax = plt.subplots(figsize=(10, 7))
    rects = ax.bar(ind,list,width,color='SkyBlue',label = '人数')
    ax.set_ylabel('人数')
    ax.set_title(chartname)
    ax.set_xticks(ind)
    ax.set_xticklabels(('优秀\n(分数≥90)', '良好\n(80≤分数<90)', '中等\n(70≤分数<80)', '及格\n(60≤分数<70)', '不及格\n(分数<60)'))
    ax.legend()
    autolabel(rects,ax)
    plt.show()

def draw_pie(list,chartname):
    labels = ['优秀\n(分数≥90)', '良好\n(80≤分数<90)', '中等\n(70≤分数<80)', '及格\n(60≤分数<70)', '不及格\n(分数<60)']
    fig1, ax1 = plt.subplots()
    explode = (0.1, 0.1, 0.1, 0.1, 0.1)
    ax1.pie(list, explode=explode, labels=labels, autopct='%1.1f%%',shadow=True, startangle=90)
    ax1.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    ax1.set_title(chartname)
    plt.show()

#分析
def analysis_lesson(list,lessonName):
    analysisWindow = tk.Toplevel()
    analysisWindow.title('Lesson analysis')
    gradeList = []
    level = [0,0,0,0,0]
    for x in list:
        a = x[gradeCol]
        gradeList.append(a)
        if a >= 90:
            level[0] += 1
        elif a >= 80:
            level[1] += 1
        elif a >= 70:
            level[2] += 1
        elif a >= 60:
            level[3] += 1
        else:
            level[4] += 1
    gradeList.sort()
    a = avg(gradeList)
    s = 0
    for x in gradeList:
        s += (x-a)**2
    std = (s/len(gradeList))**0.5
    n = int(len(gradeList)*0.27)
    topavg = avg(gradeList[len(gradeList)-n:len(gradeList)])
    nottopavg = avg(gradeList[0:n+1])
    div = (topavg-nottopavg)/100
    tk.Label(analysisWindow,text='课程成绩分析：').grid(row=0)
    tk.Label(analysisWindow,text='').grid(row=1)
    tk.Label(analysisWindow,text='考试人数：').grid(row=1,column=1)
    tk.Label(analysisWindow,text=str(len(list))).grid(row=1,column=2)
    tk.Label(analysisWindow,text='').grid(row=2)
    tk.Label(analysisWindow,text='平均分：').grid(row=2,column=1)
    tk.Label(analysisWindow,text=float('%.4f' % a)).grid(row=2,column=2)
    tk.Label(analysisWindow,text='').grid(row=3)
    tk.Label(analysisWindow,text='最高分：').grid(row=3,column=1)
    tk.Label(analysisWindow,text=str(gradeList[-1])).grid(row=3,column=2)
    tk.Label(analysisWindow,text='').grid(row=4)
    tk.Label(analysisWindow,text='最低分：').grid(row=4,column=1)
    tk.Label(analysisWindow,text=str(gradeList[0])).grid(row=4,column=2)
    tk.Label(analysisWindow,text='').grid(row=5)
    tk.Label(analysisWindow,text='标准差：').grid(row=5,column=1)
    tk.Label(analysisWindow,text=float('%.4f' % std)).grid(row=5,column=2)
    tk.Label(analysisWindow,text='').grid(row=6)
    tk.Label(analysisWindow,text='区分度：').grid(row=6,column=1)
    tk.Label(analysisWindow,text=float('%.4f' % div)).grid(row=6,column=2)
    tk.Label(analysisWindow,text='').grid(row=7)
    tk.Label(analysisWindow,text='分数段分布：').grid(row=7,column=1)
    tk.Button(analysisWindow,text='柱状图',command=lambda:draw_barchart(level,lessonName+'成绩分布柱状图')).grid(row=7,column=2)
    tk.Button(analysisWindow,text='饼图',command=lambda:draw_pie(level,lessonName+'成绩分布饼图')).grid(row=7,column=3)
    show_sheet(list,sheettitle)

def choose_analysis_lesson():
    lesson = getLesson()
    viewWindow = tk.Toplevel()
    viewWindow.title('Choose Lesson')
    tk.Button(viewWindow,text='OK!',command=lambda:analysis_lesson(lesson[viewListbox.get(viewListbox.curselection())],viewListbox.get(viewListbox.curselection()))).pack()
    listFrame = tk.Frame(viewWindow)
    scrollbar = tk.Scrollbar(listFrame, orient=tk.VERTICAL)
    viewListbox = tk.Listbox(listFrame, yscrollcommand=scrollbar.set)
    for s in lesson.keys():
        viewListbox.insert('end',s)
    scrollbar.config(command=viewListbox.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    viewListbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=1)
    listFrame.pack()

#主页面显示
def back_page(list):
    global page
    if page > 0:
        page -= 1
        show_sheet(list,sheettitle)
    else:
        tk.messagebox.showinfo(title='Page',message='Already the first page!')

def front_page(list):
    global page
    if page < pageNum-1:
        page += 1
        show_sheet(list,sheettitle)
    else:
        tk.messagebox.showinfo(title='Page',message='Already the last page!')

def jump_page(list):
    global page
    try:
        p = tkPage.get()
        if p > 0 and p < pageNum:
            page = p-1
            show_sheet(list,sheettitle)
        else:
            tk.messagebox.showinfo(title='Page',message='页数错误，请重新输入。')
    except:
        tk.messagebox.showinfo(title='Page',message='页数错误，请重新输入。')

def show_sheet(list,title):
    global tkPage,page,sheetFrame,pageNum
    tkPage = tk.IntVar()
    if len(stu)%pageSize == 0:
        pageNum = len(list)/pageSize
    else:
        pageNum = len(list)//pageSize+1
    if sheetFrame:
        sheetFrame.destroy()
    sheetFrame = tk.Frame(root)
    sheetFrame.grid(row=0)
    tkPage.set(str(page+1))
    tk.Label(sheetFrame,text='').grid(row=0)
    for i in range(len(title)):
        tk.Label(sheetFrame,text=title[i]).grid(row=0,column=i+1)
    for i in range(pageSize):
        if page*30+i >= len(list):
            break
        #tk.Label(sheetFrame,text=str(i+1),width=2).grid(row=i+1)
        tk.Label(sheetFrame,text=i+1,width=2).grid(row=i+1)
        for j in range(len(list[page*30+i])):
            if(isinstance(list[page*30+i][j],float)):
                tk.Label(sheetFrame,text=float('%.4f' % list[page*30+i][j])).grid(row=i+1,column=j+1)
            else:
                tk.Label(sheetFrame,text=list[page*30+i][j]).grid(row=i+1,column=j+1)
    statusFrame = tk.Frame(root)
    statusFrame.grid(row=1)
    tk.Button(statusFrame,text='<',command = lambda:back_page(list)).grid(row=0)
    tk.Entry(statusFrame,textvariable=tkPage,width=4).grid(row=0,column=1)
    tk.Label(statusFrame,text=' / '+str(pageNum)).grid(row=0,column=2)
    tk.Button(statusFrame,text='Go',command = lambda:jump_page(list)).grid(row=0,column=3)
    tk.Button(statusFrame,text='>',command = lambda:front_page(list)).grid(row=0,column=4)
    tk.Button(statusFrame,text='Save',command = lambda:save_sheet(list,title)).grid(row=0,column=5)

#文件访问
def click_open_button():
    global filename,stu,sheettitle
    try:
        stu = []
        filename = tkFilename.get()
        book = xlrd.open_workbook(filename)
        sheet = book.sheet_by_index(0) # 打开索引号为0的表
        for i in range(sheet.nrows):
            row = sheet.row_values(i) # 逐行读取
            if i == titleLine-1:
                sheettitle = row
            if i >= titleLine: 
                stu.append(row)
    #print(sheettitle)
        welcomeFrame.destroy()
        show_sheet(stu,sheettitle)
        openfileWindow.destroy()
        tk.messagebox.showinfo(title='Open file',message='File is opened sucessfully!')
    except:
        tk.messagebox.showwarning(title='Open file', message='打开文件错误，请检查文件路径。')

def open_file():
    global openfileWindow,tkFilename
    tkFilename = tk.StringVar()
    openfileWindow = tk.Toplevel()
    openfileWindow.title('Open file')
    openfileWindow.geometry('400x100')
    tk.Label(openfileWindow,text='Open file:').grid(row = 0)
    entryFilename = tk.Entry(openfileWindow,textvariable = tkFilename)
    entryFilename.grid(row = 0,column = 1)
    tk.Button(openfileWindow,text='Browse',command=browse_file).grid(row = 0,column = 2)
    tk.Button(openfileWindow,text='Open',command=click_open_button).grid(row = 1,column = 1)

def save_sheet(table,title):
    book = xlwt.Workbook(encoding = 'utf-8')
    sheet = book.add_sheet("sheet1")
    for i in range(len(title)):
        sheet.write(0,i,title[i])
    for i in range(len(table)):
        for j in range(len(table[i])):
            sheet.write(i+1,j,table[i][j])
    try:
        filename = filedialog.asksaveasfilename(filetypes=[('Excel file','.xlsx')])
        print(filename)
        book.save(filename)
    except:
        tk.messagebox.showwarning(title='Close file', message='保存文件错误，请检查文件路径。')

#排名
def rank_lesson():
    lesson = getLesson()
    sortedlesson = {}
    for x in lesson:
        sortedlesson[x]=sorted(lesson[x],key=lambda s :s[gradeCol],reverse=True)
    show_view(sortedlesson)

#GPA
def normalGPA(list):
    ans = []
    for x in list.values():
        outtemplist = []
        for y in x:
            templist = []
            if y[1] >= 90:
                templist = [y[0],4.0]
            elif y[1] >= 80:
                templist = [y[0],3.0]
            elif y[1] >= 70:
                templist = [y[0],2.0]
            elif y[1] >= 60:
                templist = [y[0],1.0]
            else:
                templist = [y[0],0]
            outtemplist.append(templist)
        ans.append(outtemplist)
    #print(ans)
    return weightavg(ans)

def normalWeightavg(list):
    return(weightavg(list.values()))

def analysis_GPA(lessonName,lessonCheckList):
    global page
    lessonListWindow.destroy()
    GPALesson = []
    for i in range(len(lessonCheckList)):
        if lessonCheckList[i].get() == 1:
            GPALesson.append(lessonName[i])
    GPACalculator = {}
    #获取每个学生所要计算课程的学分和成绩
    for x in stu:
        keyname = x[studentIdCol]+' '+x[studentNameCol]
        s = x[lessonIdCol]+' '+x[lessonNameCol]
        if s in GPALesson:
            y = []
            y.append(x[creditCol])
            y.append(x[gradeCol])
            if keyname not in GPACalculator.keys():
                GPACalculator[keyname] = [y]
            else:
                GPACalculator[keyname].append(y)
    #计算每个学生的各种绩点
    GPA = []
    GPA.append(GPACalculator.keys())
    GPA.append(normalWeightavg(GPACalculator))
    GPA.append(normalGPA(GPACalculator))
    title = ['','加权平均数','标准GPA']
    GPA = list(map(list,zip(*GPA)))
    page = 0
    show_sheet(GPA,title)

def choose_GPA_lesson():
    global lessonListWindow
    lessonName = getLessonName(stu)
    lessonListWindow = tk.Toplevel()
    lessonListWindow.title('选择课程')
    tk.Label(lessonListWindow,text='请选择需要计入总成绩的课程：')
    lessonCheckList = [tk.IntVar() for i in range(len(lessonName))]
    for i in range(len(lessonName)):
        tk.Checkbutton(lessonListWindow, text=lessonName[i], variable=lessonCheckList[i], onvalue=1, offvalue=0).pack()
    tk.Button(lessonListWindow,text='确定',command =lambda:analysis_GPA(lessonName,lessonCheckList)).pack()

#主界面
root = tk.Tk()
root.title('Grade Analysis System')
root.geometry('1280x720')

menu = tk.Menu(root)
root.config(menu=menu)

filemenu = tk.Menu(menu)
filemenu.add_command(label='Open',command=open_file)
filemenu.add_command(label='Exit',command=exit_program)
menu.add_cascade(label='File',menu=filemenu)

setmenu = tk.Menu(menu)
setmenu.add_command(label='System Variable',command=set_variable)
menu.add_cascade(label='Set',menu=setmenu)

viewmenu = tk.Menu(menu)
viewmenu.add_command(label='Lesson',command=show_lesson)
viewmenu.add_command(label='Student',command=show_student)
menu.add_cascade(label='View',menu=viewmenu)

rankmenu = tk.Menu(menu)
rankmenu.add_command(label='Lesson',command=rank_lesson)
menu.add_cascade(label='Rank',menu=rankmenu)

analysismenu = tk.Menu(menu)
analysismenu.add_command(label='Lesson',command=choose_analysis_lesson)
analysismenu.add_command(label='GPA',command=choose_GPA_lesson)
menu.add_cascade(label='Analysis',menu=analysismenu)

helpmenu = tk.Menu(menu)
helpmenu.add_command(label='About',command=show_about)
menu.add_cascade(label='Help',menu=helpmenu)

welcomeFrame = tk.Frame(root)
welcomeFrame.pack()
tk.Label(welcomeFrame,text='Welcome to use my Grade Analysis System.').pack()
tk.Label(welcomeFrame,text='Let\'s click File >> Open to start!.').pack()

root.mainloop()