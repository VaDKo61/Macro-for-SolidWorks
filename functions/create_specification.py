import os

import pythoncom
import win32com.client


def create_specification(sw_app, sw_model, vt_dispatch):
    assembly_path: str = sw_model.GetPathName
    assembly_path_list: list = assembly_path.split('\\')[2:]
    assembly_name: str = assembly_path_list[-1].split('.')[0]
    engineer: str = assembly_path_list[2]

    # select frame part
    frames: list = []
    for component in sw_model.GetComponents(True):
        if 'Рама' in component.Name2:
            frames.append(component.GetPathName)

    # create draw
    template_path = f'\\\\{assembly_path_list[0]}\\{assembly_path_list[1]}\\{engineer}\\Шаблоны\\Чертеж сборки.DRWDOT'
    sw_draw = sw_app.NewDocument(template_path, 12, 0.42, 0.297)

    # add draw view
    x_view: float = 0
    y_view: float = 0.148
    frames_view: list = []
    assembly_view = sw_draw.CreateDrawViewFromModelView3(assembly_path, '*Изометрия', x_view, y_view, 0)
    for frame in frames:
        x_view += 0.42
        frames_view.append(sw_draw.CreateDrawViewFromModelView3(frame, '*Изометрия', x_view, y_view, 0))

    # add specification
    template_assembly = '\\\\192.168.1.14\\SolidWorks\\Библиотека Solid Works НОВАЯ\\Шаблоны\\' \
                        'Шаблон специи.sldbomtbt'
    template_frame = '\\\\192.168.1.14\\SolidWorks\\Библиотека Solid Works НОВАЯ\\Шаблоны\\' \
                     'Сварные конструкции.sldwldtbt'
    x_spec: float = -0.2
    y_spec: float = -0.03
    table_assembly = assembly_view.InsertBomTable4(False, x_spec, y_spec, 1, 1, 'По умолчанию', template_assembly,
                                                   False, 0, False)
    tables_frame: list = []
    for view in frames_view:
        x_spec += 0.42
        tables_frame.append(
            view.InsertWeldmentTable(False, x_spec, y_spec, 1, 'По умолчанию<Как сварной>', template_frame))

    # save draw
    path_specification: str = '\\'.join(assembly_path.split('\\')[:-1]) + '\\Спецификации'
    if not os.path.isdir(path_specification):
        os.makedirs(path_specification)
    path_draw = f'{path_specification}\\{assembly_name} (Cпецификация).SLDDRW'
    sw_draw.SaveAs3(path_draw, 0, 0)

    # save table excel
    table_assembly.SaveAsExcel(f'{path_specification}//{assembly_name}.xlsx', False, False)
    for i, table in enumerate(tables_frame, start=1):
        text = table.WeldmentCutListFeature.GetTableAnnotations[0]
        text.SaveAsText2(f'{path_specification}//{assembly_name} (Рама ч{i}).txt', '__', False)
    


def specification():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    vt_dispatch = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 2:
        sw_app.SendmsgToUser('Активна не сборка')
        print('Активна не сборка')
        return
    create_specification(sw_app, sw_model, vt_dispatch)


specification()
