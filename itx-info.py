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

    def update(self,event=None):
        # Check if necessary folder exists
        if not os.path.isdir('Keep Out'):
            os.mkdir('Keep Out')

        os.chdir('Keep Out')
        
        # Check if Excel file exists
        if not os.path.isfile('Master.xlsx'):
            workbook = xl.Workbook()
            worksheet = workbook.get_active_sheet()
            worksheet.title = 'ITX'
            headers = ['ITX','A-Version','A-Hardware','B-Version','B-Hardware']
            for i in range(5):
                worksheet.cell(row=0,column=i).value = headers[i]
                
            workbook.save(filename='Master.xlsx')

        workbook = xl.load_workbook(filename = 'Master.xlsx')
        worksheet = workbook.get_sheet_by_name(name = 'ITX')

        itx = str(self.inputs[0].get())
        if itx == '':
            os.chdir('..')
            return
        if string.lower(itx)[-1] == 'a':
            self.writeline(1,2,workbook,itx)
                
        elif string.lower(itx)[-1] == 'b':
            self.writeline(3,4,workbook,itx)

        else:
            print 'Invalid input: ITX'
            
        os.chdir('..')

    # Handle main or backup cases
    def writeline(self,a,b,workbook,itx):
        worksheet = workbook.get_sheet_by_name(name = 'ITX')
        i = 1
        exists = False
        while worksheet.cell(row=i,column=0).value != None:

            # Checks if ITX entry exists
            if str(worksheet.cell(row=i,column=0).value) == itx[:-1]:

                # Opens popup to confirm overwrite
                if worksheet.cell(row=i,column=a).value != None or \
                   worksheet.cell(row=i,column=b).value != None:

                    popupstring = 'Existing entry:\nITX:'+\
                            string.upper(itx)+' V: '+\
                            str(worksheet.cell(row=i,column=a).value)+' H: '+\
                            str(worksheet.cell(row=i,column=b).value)+'.'+\
                            '\n\nOverwrite?\n'
                    inputDialog = Popup(root,popupstring)
                    root.wait_window(inputDialog.top)

                try:
                    if overwrite:
                        worksheet.cell(row=i,column=a).value = \
                                str(self.inputs[1].get())
                        worksheet.cell(row=i,column=b).value = \
                                str(self.inputs[2].get())
                        popupstring = 'Successfully entered:\n\nITX:'+\
                            string.upper(itx)+' V: '+\
                            str(worksheet.cell(row=i,column=a).value)+' H: '+\
                            str(worksheet.cell(row=i,column=b).value)+'.\n'
                        successDialog = Success(root,popupstring)
                        root.wait_window(successDialog.top)
                except Exception:
                    ''
                exists = True
                break
            i += 1

        if not exists:
            worksheet.cell(row=i,column=0).value = \
                            str(self.inputs[0].get())[:-1]
            worksheet.cell(row=i,column=a).value = \
                            str(self.inputs[1].get())
            worksheet.cell(row=i,column=b).value = \
                            str(self.inputs[2].get())
            popupstring = 'Successfully entered:\n\nITX:'+\
                            string.upper(itx)+' V: '+\
                            str(worksheet.cell(row=i,column=a).value)+' H: '+\
                            str(worksheet.cell(row=i,column=b).value)+'.\n'
            successDialog = Success(root,popupstring)
            root.wait_window(successDialog.top)
        workbook.save(filename='Master.xlsx')


        
class Popup:
    def __init__(self, parent, popupstring):
        top = self.top = tk.Toplevel(parent)
        self.myLabel = tk.Label(top, text=popupstring)
        self.myLabel.grid(column=0,row=0,columnspan=2)

        self.yesButton = tk.Button(top, text='Yes',
                                   command=lambda : self.send(True))
        self.yesButton.grid(row=1,column=0,sticky=tk.E)
        self.noButton = tk.Button(top, text='No',
                                  command=lambda : self.send(False))
        self.noButton.grid(row=1,column=1,sticky=tk.W)

    def send(self,doit):
        global overwrite
        overwrite = doit
        self.top.destroy()

class Success:
    def __init__(self, parent, popupstring):
        top = self.top = tk.Toplevel(parent)
        self.myLabel = tk.Label(top, text=popupstring)
        self.myLabel.grid(column=0,row=0)

        self.closeButton = tk.Button(top, text='Close',
                                   command=self.send)
        self.closeButton.grid(row=1,column=0)

    def send(self):
        self.top.destroy()
    
# MAIN PROCESS
root = tk.Tk()
#root.geometry('800x600')
#root.resizable(0,0)
#root.minsize(width=550,height=450)
app = Application()
#root.bind("<Return>",app.update)
app.master.title('ITX Info')

app.mainloop()
