from tkinter import Tk, Button, Label

from conversion_excel import conversion_excel

root = Tk()
root.title('Помощник инженера')
root.geometry('600x600')

label = Label(text='Функции для инженера-конструктора', font=('Times New Roman', 15))
label.pack()

btn_conversion_excel = Button(root,
                              text='Преобразовать Excel',
                              command=conversion_excel,
                              font=('Times New Roman', 15),
                              activebackground='red')
btn_conversion_excel.place(x=10, y=40)

root.mainloop()
