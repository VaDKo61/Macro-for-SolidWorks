import os

import pythoncom
import win32com.client

from functions.create_drill_sheet import get_ready


def create_path_tube(path: str):
    """Check and crete path"""
    try:
        os.makedirs(path)
        print(f'Директория {path} была создана')
    except FileExistsError:
        print(f'Директория {path} уже существует')


def save_tube():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    arg1 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 2)
    arg2 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128)
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 2:
        sw_app.SendmsgToUser('Активна не сборка')
        print('Активна не сборка')
        return

    # save tube in a separate file
    tubes: list = []
    if get_ready():
        return
    for component in sw_model.GetComponents(True):
        component_name: str = component.name2.split('-')[0]
        if component_name.startswith(('Труба', 'Ниппель', 'Резьба')):
            if component_name not in tubes:
                tubes.append(component_name)
                part = component.GetModelDoc2
                path_list: list = sw_model.GetPathName.split('\\')
                assembly_name: str = path_list.pop().split('.')[0]
                path_list.append('Трубы')
                path_list.append(assembly_name)
                path: str = '\\'.join(path_list)
                create_path_tube(path)
                part.SaveAs3(f'{path}\\{component_name}.SLDPRT', 0, 8)
    else:
        sw_app.SendmsgToUser('Трубы успешно сохранены')
        print('Трубы успешно сохранены')
    sw_model.Save3(1, arg1, arg2)
