from tkinter import Tk, Button, Label

from functions.add_length_tube import length_tube
from functions.conversion_excel import conversion_excel
from functions.create_any_cut_extrude import any_cut_extrude, any_cut_extrude_kip
from functions.create_cut_extrude import cut_extrude, cut_extrude_kip
from functions.create_drawing import drawing
from functions.create_saddle_assembly import assembly_saddle
from functions.create_saddle_part import part_saddle
from functions.create_specification import specification
from functions.save_elements_frame_igs import elements_frame_igs
from functions.save_igs import save_igs
from functions.save_tube import save_tube

root = Tk()
root.title('Помощник инженера')
root.geometry('700x600')

label_1 = Label(text='Чертежи:', font=('Times New Roman', 15))
label_1.place(x=10, y=10)
label_2 = Label(text='Спецификации:', font=('Times New Roman', 15))
label_2.place(x=10, y=80)
label_3 = Label(text='Лазерный станок:', font=('Times New Roman', 15))
label_3.place(x=10, y=150)

btn_drawing = Button(root,
                     text='Создать чертеж (1-5 лист)',
                     command=drawing,
                     font=('Times New Roman', 13),
                     activebackground='red',
                     cursor="hand2")
btn_drawing.place(x=10, y=40)
btn_length_tube = Button(root,
                         text='Добавить размеры труб на вид',
                         command=length_tube,
                         font=('Times New Roman', 13),
                         activebackground='red',
                         cursor="hand2")
btn_length_tube.place(x=215, y=40)
btn_specification = Button(root,
                           text='Создать спецификацию (Чертеж)',
                           command=specification,
                           font=('Times New Roman', 13),
                           activebackground='red',
                           cursor="hand2")
btn_specification.place(x=10, y=110)
btn_conversion_excel = Button(root,
                              text='Преобразовать Excel',
                              command=conversion_excel,
                              font=('Times New Roman', 13),
                              activebackground='red',
                              cursor="hand2")
btn_conversion_excel.place(x=261, y=110)
btn_save_tube = Button(root,
                       text='Сохранить все трубы в папку проекта',
                       command=save_tube,
                       font=('Times New Roman', 13),
                       activebackground='red',
                       cursor="hand2")
btn_save_tube.place(x=10, y=180)
btn_create_saddle = Button(root,
                           text='Создать седло в сборке',
                           command=assembly_saddle,
                           font=('Times New Roman', 13),
                           activebackground='red',
                           cursor="hand2")
btn_create_saddle.place(x=10, y=220)
btn_create_saddle = Button(root,
                           text='Создать седло в детали',
                           command=part_saddle,
                           font=('Times New Roman', 13),
                           activebackground='red',
                           cursor="hand2")
btn_create_saddle.place(x=198, y=220)
btn_cut_extrude = Button(root,
                         text='Создать отверстие от трубы',
                         command=cut_extrude,
                         font=('Times New Roman', 13),
                         activebackground='red',
                         cursor="hand2")
btn_cut_extrude.place(x=10, y=260)
btn_cut_extrude_kip = Button(root,
                             text='Создать отверстие от врезки (КИП)',
                             command=cut_extrude_kip,
                             font=('Times New Roman', 13),
                             activebackground='red',
                             cursor="hand2")
btn_cut_extrude_kip.place(x=228, y=260)
btn_any_cut_extrude = Button(root,
                             text='Создать отверстие от нескольких труб',
                             command=any_cut_extrude,
                             font=('Times New Roman', 13),
                             activebackground='red',
                             cursor="hand2")
btn_any_cut_extrude.place(x=10, y=300)
btn_any_cut_extrude_kip = Button(root,
                                 text='Создать отверстие от нескольких врезок (КИП)',
                                 command=any_cut_extrude_kip,
                                 font=('Times New Roman', 13),
                                 activebackground='red',
                                 cursor="hand2")
btn_any_cut_extrude_kip.place(x=301, y=300)
btn_save_igs = Button(root,
                      text='Сохранить все активные трубы в IGS',
                      command=save_igs,
                      font=('Times New Roman', 13),
                      activebackground='red',
                      cursor="hand2")
btn_save_igs.place(x=10, y=340)
btn_save_elements_frame_igs = Button(root,
                                     text='Сохранить элементы рамы в IGS',
                                     command=elements_frame_igs,
                                     font=('Times New Roman', 13),
                                     activebackground='red',
                                     cursor="hand2")
btn_save_elements_frame_igs.place(x=10, y=380)

root.mainloop()
