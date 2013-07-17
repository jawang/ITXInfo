import openpyxl as xl
import Tkinter as tk
import os
import string

class Application(tk.Frame):
    def __init__(self,master=None):

        tk.Frame.__init__(self,master)

        self.grid()
        self.createWidgets()

    def createWidgets(self):
        # Inputs #################################
        self.inputText = ['ITX ID', 'Version', 'Hardware']
        self.inputs = [tk.StringVar() for i in range(3)]
        self.inputLab = [tk.Label(self,text=self.inputText[i])
                         for i in range(3)]
        self.inputBox = [tk.Entry(self,textvariable=self.inputs[i])
                         for i in range(3)]

        for i in range(3):
            self.inputLab[i].grid(row=0,column=i)
            self.inputBox[i].grid(row=1,column=i)

        # Update button ##########################
        self.updateButton = tk.Button(command=self.update,text='Update')
        self.updateButton.grid(row=2,column=0,columnspan=3)

    def update(self):
        # Check if necessary folder exists
        if not os.path.isdir('Keep Out'):
            os.mkdir('Keep Out')

        os.chdir('Keep Out')
        
        # Check if Excel file exists
        if not os.path.isfile('Master.xlsx'):
            workbook = xl.Workbook()
            worksheet = workbook.create_sheet()
            worksheet.title = 'ITX'
            headers = ['ITX','A-Version','A-Hardware','B-Version','B-Hardware']
            for i in range(5):
                worksheet.cell(row=0,column=i).value = headers[i]
                
            workbook.save(filename='Master.xlsx')

        workbook = xl.load_workbook(filename = 'Master.xlsx')
        worksheet = workbook.get_sheet_by_name(name = 'ITX')

        itx = str(self.inputs[0].get())
        if string.lower(itx)[-1] == 'a':
            print 'asdf'
            i = 1
            exists = False
            while worksheet.cell(row=i,column=0).value != None:
                if str(worksheet.cell(row=i,column=0).value) == itx[:-1]:
                    worksheet.cell(row=i,column=1).value = \
                                str(self.inputs[1].get())
                    worksheet.cell(row=i,column=2).value = \
                                str(self.inputs[2].get())
                    exists = True
                    break
                i += 1

            if not exists:
                worksheet.cell(row=i,column=0).value = \
                                str(self.inputs[0].get())[:-1]
                worksheet.cell(row=i,column=1).value = \
                                str(self.inputs[1].get())
                worksheet.cell(row=i,column=2).value = \
                                str(self.inputs[2].get())

            workbook.save(filename='Master.xlsx')
                
        elif string.lower(itx)[-1] == 'b':
            i = 1
            exists = False
            while worksheet.cell(row=i,column=0).value != None:
                if str(worksheet.cell(row=i,column=0).value) == itx[:-1]:
                    worksheet.cell(row=i,column=3).value = \
                                str(self.inputs[1].get())
                    worksheet.cell(row=i,column=4).value = \
                                str(self.inputs[2].get())
                    exists = True
                    break
                i += 1

            if not exists:
                worksheet.cell(row=i,column=0).value = \
                                str(self.inputs[0].get())[:-1]
                worksheet.cell(row=i,column=3).value = \
                                str(self.inputs[1].get())
                worksheet.cell(row=i,column=4).value = \
                                str(self.inputs[2].get())
        
            workbook.save(filename='Master.xlsx')

        else:
            print 'Invalid input: ITX'
            
        os.chdir('..')

class Popup:
    def __init__(self,parent):
        ''
        
        
    
# MAIN PROCESS
root = tk.Tk()
#root.geometry('800x600')
#root.resizable(0,0)
#root.minsize(width=550,height=450)
app = Application()
#root.bind("<Configure>",app.resize)
app.master.title('ITX Info')

app.mainloop()
