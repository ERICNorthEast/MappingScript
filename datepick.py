from Tkinter import *
from ttk import *
import datetime
from calendar import monthrange

class DatePicker(Frame):

    def createDayButtons(self):
        currentMonthRange = monthrange(self.dateValue.year, self.dateValue.month)

        if self.dateValue.month == datetime.date.today().month:
            currentMonth = True
        else:
            currentMonth = False

        self.title['text']=self.dateValue.strftime('%B %Y')

        self.dayButtons = []
        for i in range(1, (currentMonthRange[1]+1)):
            dayButton = Button(self, text=str(i), width=2, command=self.getDay)
            if currentMonth and i == datetime.date.today().day:
                dayButton.configure(style='today.TButton')
            self.dayButtons.append(dayButton)
        
        currentRow=5
        i=currentMonthRange[0]
        for dayButton in self.dayButtons:
            dayButton.grid(row=currentRow, column=i)
            i+=1
            if i>6:
                i=0
                currentRow+=1
                

    def createWidgets(self):
        self.btnPrev = Button(self, text='<', command=self.prevMonth, width=2)
        self.title = Label(self)
        self.btnNext = Button(self, text='>', command=self.nextMonth, width=2)
        days=['Mo','Tu','We','Th','Fr','Sa','Su']
        self.dayTitles = []
        for day in days:
            title = Label(self, text=day)
            self.dayTitles.append(title)

        self.btnPrev.grid(row=0, column=0)
        self.title.grid(row=0, column=1, columnspan=5)
        self.btnNext.grid(row=0, column=6)

        i = 0
        for dayTitle in self.dayTitles:
            dayTitle.grid(row=3, column=i)
            i+=1

        self.createDayButtons()

        self.btnCancel = Button(self, text='Cancel', command=self.cancel, width=8)
        self.btnCancel.grid(row=12, column=3, columnspan=4, sticky=E)


    def clearDays(self):
        for dayButton in self.dayButtons:
            dayButton.destroy()


    def clearAll(self):
        self.btnPrev.destroy()
        self.title.destroy()
        self.btnNext.destroy()
        for dayTitle in self.dayTitles:
            dayTitle.destroy()
        self.btnCancel.destroy()
        self.clearDays()
        
                
    def prevMonth(self):
        self.clearDays()
        self.dateValue -= datetime.timedelta(days=1)
        self.dateValue = self.dateValue.replace(day=1)
        print("new date %s" % self.dateValue)
        self.createDayButtons()


    def nextMonth(self):
        self.clearDays()
        self.dateValue += datetime.timedelta(days=31)
        self.dateValue = self.dateValue.replace(day=1)
        print("new date %s" % self.dateValue)
        self.createDayButtons()


    def getDay(self):
        selDay = int(self.focus_get().cget('text'))
        self.returnDate = self.dateValue.replace(day=selDay)
        print self.returnDate
        #self.clearAll()
        self.app.newDate(self.returnDate)
        


    def cancel(self):
        print("nothing picked")
        #self.returnDate = self.dateValue.replace(year=1970)
        #self.clearAll()
        self.app.newDate(self.returnDate)
        
    

    def __init__(self, master, app):
        Frame.__init__(self, master)
        self.app = app
        self.dateValue = datetime.date.today().replace(day=1)
        self.returnDate = None
        Style().configure('today.TButton', background='red')
        print (self.dateValue)
        self.createWidgets()
        
