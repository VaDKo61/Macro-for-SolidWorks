import os

import pythoncom
import win32com.client


def save_elements_frame_igs(sw_app, sw_model, vt_dispatch, arg1, arg2):
    """save_elements_frame"""
    # selection manager
    selection_manager = sw_model.SelectionManager
    selection_data = selection_manager.CreateSelectData

    # create path frame
    frame_path_list = sw_model.GetPathName.split('\\')
    assembly_name = frame_path_list[-1].split('.')[0]
    frame_path = '\\'.join(frame_path_list[:-1]) + '\\Лазер\\Трубы\\' + ' '.join(assembly_name.split(' ')[0:2]) \
                 + ' Лазер' + f'\\{" ".join(assembly_name.split(" ")[2:])} IGS'
    create_path_frame(frame_path)

    sw_model.Extension.SelectByID2('Твердые тела', 'BDYFOLDER', 0, 0, 0, False, 0, vt_dispatch, 0)
    # sub_body_folder = selection_manager.GetSelectedObject6(1, -1).GetSpecificFeature2
    bodies = selection_manager.GetSelectedObject6(1, -1)
    bodies.GetSpecificFeature2.SetAutomaticCutList(True)
    bodies.GetSpecificFeature2.SetAutomaticUpdate(True)
    sw_model.ClearSelection2(True)
    bodies = bodies.GetFirstSubFeature
    arg3 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, None)
    arg4 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BSTR, None)
    arg5 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BOOL, True)
    arg6 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_BOOL, True)
    while True:
        if bodies is None:
            break
        bodies_count = bodies.GetSpecificFeature2.GetBodyCount
        if not bodies_count:
            bodies = bodies.GetNextSubFeature
            continue
        body = bodies.GetSpecificFeature2.GetBodies[0]
        bodies.CustomPropertyManager.Get6('Длина', False, arg3, arg4, arg5, arg6)
        body.Select2(True, selection_data)
        sw_model.SaveToFile3(f'{frame_path}\\{bodies.Name.replace("<", "(").replace(">", ")")} l={arg4.value} мм '
                             f'({bodies_count} шт).IGS',
                             2, 2, False, False, arg1, arg2)
        sw_app.CloseDoc('')
        sw_model.ClearSelection2(True)
        bodies = bodies.GetNextSubFeature


def create_path_frame(path: str):
    """Check and crete path"""
    try:
        os.makedirs(path)
        print(f'Директория {path} была создана')
    except FileExistsError:
        print(f'Директория {path} уже существует')


def elements_frame_igs():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    vt_dispatch = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    arg1 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 2)
    arg2 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128)
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 1:
        sw_app.SendmsgToUser('Активна не деталь')
        print('Активна не деталь')
        return
    save_elements_frame_igs(sw_app, sw_model, vt_dispatch, arg1, arg2)
    sw_app.SendmsgToUser('Элементы рамы успешно сохранены')
    print('Элементы рамы успешно сохранены')
