from tkinter import Tk, Button, Label

from functions.conversion_excel import conversion_excel
from functions.create_cut import main_any_cut
from functions.create_drawing import drawing
from functions.create_select_conf import main_create_select_conf
from functions.create_specification import specification
from functions.save_elements_frame_igs import elements_frame_igs
from functions.save_igs import main_save_igs
from functions.save_pipe import main_save_pipe

root = Tk()
root.title('Помощник инженера')
root.geometry('750x600')

label_1 = Label(text='Чертежи:', font=('Times New Roman', 15))
label_1.place(x=10, y=10)
label_2 = Label(text='Спецификации:', font=('Times New Roman', 15))
label_2.place(x=10, y=80)
label_3 = Label(text='Лазерный станок:', font=('Times New Roman', 15))
label_3.place(x=10, y=150)

btn_drawing = Button(root,
                     text='Создать чертеж (2-5 лист)',
                     command=drawing,
                     font=('Times New Roman', 13),
                     activebackground='red',
                     cursor="hand2")
btn_drawing.place(x=10, y=40)
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
                       command=main_save_pipe,
                       font=('Times New Roman', 13),
                       activebackground='red',
                       cursor="hand2")
btn_save_tube.place(x=10, y=180)
btn_create_select_conf = Button(root,
                                text='Создать конфигурации выбранных труб',
                                command=main_create_select_conf,
                                font=('Times New Roman', 13),
                                activebackground='red',
                                cursor="hand2")
btn_create_select_conf.place(x=10, y=220)
btn_any_cut_extrude = Button(root,
                             text='Создать отверстия от труб',
                             command=main_any_cut,
                             font=('Times New Roman', 13),
                             activebackground='red',
                             cursor="hand2")
btn_any_cut_extrude.place(x=10, y=260)
btn_save_igs = Button(root,
                      text='Сохранить все активные трубы в IGS',
                      command=main_save_igs,
                      font=('Times New Roman', 13),
                      activebackground='red',
                      cursor="hand2")
btn_save_igs.place(x=10, y=300)
btn_save_elements_frame_igs = Button(root,
                                     text='Сохранить элементы рамы в IGS',
                                     command=elements_frame_igs,
                                     font=('Times New Roman', 13),
                                     activebackground='red',
                                     cursor="hand2")
btn_save_elements_frame_igs.place(x=10, y=340)
# btn_create_saddle_front = Button(root,
#                                  text='Создать седло в сборке (Спереди)',
#                                  command=assembly_saddle_front,
#                                  font=('Times New Roman', 13),
#                                  activebackground='red',
#                                  cursor="hand2")
# btn_create_saddle_front.place(x=10, y=490)
# btn_create_saddle_above = Button(root,
#                                  text='Создать седло в сборке (Сверху)',
#                                  command=assembly_saddle_above,
#                                  font=('Times New Roman', 13),
#                                  activebackground='red',
#                                  cursor="hand2")
# btn_create_saddle_above.place(x=273, y=490)
# btn_create_saddle = Button(root,
#                            text='Создать седло в детали',
#                            command=part_saddle,
#                            font=('Times New Roman', 13),
#                            activebackground='red',
#                            cursor="hand2")
# btn_create_saddle.place(x=527, y=490)

root.mainloop()
