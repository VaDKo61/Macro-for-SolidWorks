from tkinter import Tk, Button, Label

from functions.conversion_excel import conversion_excel
from functions.create_cut_extrude import cut_extrude, cut_extrude_kip
from functions.create_saddle_assembly import assembly_saddle
from functions.create_saddle_part import part_saddle
from functions.create_specification import specification
from functions.save_igs import save_igs
from functions.save_tube import save_tube

root = Tk()
root.title('Помощник инженера')
root.geometry('600x600')

label_1 = Label(text='Для спецификаций:', font=('Times New Roman', 15))
label_1.place(x=10, y=10)
label_2 = Label(text='Для лазерного станка:', font=('Times New Roman', 15))
label_2.place(x=10, y=80)
label_3 = Label(text='Создание "седел":', font=('Times New Roman', 15))
label_3.place(x=10, y=150)
label_4 = Label(text='Создание отверстий:', font=('Times New Roman', 15))
label_4.place(x=10, y=220)
label_5 = Label(text='Чертежи:', font=('Times New Roman', 15))
label_5.place(x=10, y=290)

btn_conversion_excel = Button(root,
                              text='Преобразовать Excel',
                              command=conversion_excel,
                              font=('Times New Roman', 13),
                              activebackground='red',
                              cursor="hand2")
btn_conversion_excel.place(x=10, y=40)
btn_save_tube = Button(root,
                       text='Сохранить все трубы в папку проекта',
                       command=save_tube,
                       font=('Times New Roman', 13),
                       activebackground='red',
                       cursor="hand2")
btn_save_tube.place(x=10, y=110)
btn_save_igs = Button(root,
                      text='Сохранить все активные трубы в IGS',
                      command=save_igs,
                      font=('Times New Roman', 13),
                      activebackground='red',
                      cursor="hand2")
btn_save_igs.place(x=300, y=110)
btn_create_saddle = Button(root,
                           text='Создать седло в сборке',
                           command=assembly_saddle,
                           font=('Times New Roman', 13),
                           activebackground='red',
                           cursor="hand2")
btn_create_saddle.place(x=10, y=180)
btn_create_saddle = Button(root,
                           text='Создать седло в детали',
                           command=part_saddle,
                           font=('Times New Roman', 13),
                           activebackground='red',
                           cursor="hand2")
btn_create_saddle.place(x=200, y=180)
btn_cut_extrude = Button(root,
                         text='Создать отверстие от трубы',
                         command=cut_extrude,
                         font=('Times New Roman', 13),
                         activebackground='red',
                         cursor="hand2")
btn_cut_extrude.place(x=10, y=250)
btn_cut_extrude_kip = Button(root,
                             text='Создать отверстие от врезки (КИП)',
                             command=cut_extrude_kip,
                             font=('Times New Roman', 13),
                             activebackground='red',
                             cursor="hand2")
btn_cut_extrude_kip.place(x=230, y=250)
btn_specification = Button(root,
                           text='Создать спецификацию',
                           command=specification,
                           font=('Times New Roman', 13),
                           activebackground='red',
                           cursor="hand2")
btn_specification.place(x=10, y=320)

root.mainloop()
