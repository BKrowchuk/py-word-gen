from tkinter import font
from docxtpl import DocxTemplate
from tkinter import *


def submitPressed():
    # Document Creation
    doc_title=dTitle.get()
    doc = DocxTemplate("word_template.docx")
    context = { 'name' : dName.get(),
    'location': dLocation.get(),
    'ip': dIP.get() }
    doc.render(context)
    doc.save(doc_title + ".docx")
    window.destroy()

wHeight = 325
wWidth = 300
xVal = 100

# Tkinter form
window=Tk()

# labels
# lblTitle=Label(window, text="Input Form", font=("Helvetica", 16))
# lblTitle.place(x=50, y=50)
lblTitle=Label(window, text="Input Title:")
lblTitle.place(x=xVal, y=wHeight-295)
lblName=Label(window, text="Input Name:")
lblName.place(x=xVal, y=wHeight-235)
lblLocation=Label(window, text="Input Location:")
lblLocation.place(x=xVal, y=wHeight-175)
lblIP=Label(window, text="Input IP:")
lblIP.place(x=xVal, y=wHeight-115)

# variables
dTitle = StringVar()
dName = StringVar()
dLocation = StringVar()
dIP = StringVar()

# inputs
txtName=Entry(window, bd=2, textvariable=dTitle)
txtName.place(x=xVal, y=wHeight-270)
txtName=Entry(window, bd=2, textvariable=dName)
txtName.place(x=xVal, y=wHeight-210)
txtLocation=Entry(window, bd=2, textvariable=dLocation)
txtLocation.place(x=xVal, y=wHeight-150)
txtIP=Entry(window, bd=2, textvariable=dIP)
txtIP.place(x=xVal, y=wHeight-90)

# buttons
btnSubmit=Button(window, text="Submit", command=submitPressed)
btnSubmit.place(x=wWidth-100, y=wHeight-50)
btnCancel=Button(window, text="Cancel", command=window.destroy)
btnCancel.place(x=xVal, y=wHeight-50)

# window information
window.title('MOP Generation')
window.geometry("350x350+10+10")
window.mainloop()

