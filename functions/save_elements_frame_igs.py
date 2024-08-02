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
    frame_path = '\\'.join(frame_path_list[:-1]) + '\\Рама\\' + frame_path_list[-1].split('.')[0]
    create_path_frame(frame_path)

    sw_model.Extension.SelectByID2('Твердые тела', 'BDYFOLDER', 0, 0, 0, False, 0, vt_dispatch, 0)
    # sub_body_folder = selection_manager.GetSelectedObject6(1, -1).GetSpecificFeature2
    bodies = selection_manager.GetSelectedObject6(1, -1)
    bodies.GetSpecificFeature2.SetAutomaticCutList(True)
    bodies.GetSpecificFeature2.SetAutomaticUpdate(True)
    sw_model.ClearSelection2(True)
    while True:
        bodies = bodies.GetNextFeature
        if bodies.GetSpecificFeature2 is None:
            break
        if not bodies.GetSpecificFeature2.GetBodyCount:
            continue
        for body in bodies.GetSpecificFeature2.GetBodies:
            body.Select2(True, selection_data)
            sw_model.SaveToFile3(f'{frame_path}\\{body.Name}.IGS', 2, 2, False, False,
                                 arg1, arg2)
            sw_app.CloseDoc('')
            sw_model.ClearSelection2(True)


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
