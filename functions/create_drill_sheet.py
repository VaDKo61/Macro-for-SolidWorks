import os

import pythoncom
import win32com.client


def create_drill_sheet(sw_app, sw_model, vt_dispatch, arg1, arg2):
    """create view all elements frame where is drill"""
    selection_manager = sw_model.SelectionManager
    frame_view = selection_manager.GetSelectedObject6(1, -1)
    path_model_frame = frame_view.GetReferencedModelName
    path_draw = sw_model.GetPathName
    sw_app.CloseDoc(path_draw.split('\\')[-1])
    sw_app.OpenDoc6(path_model_frame, 1, 2, '', arg1, arg2)

    # create path frame
    sw_model = sw_app.ActiveDoc
    selection_manager_part = sw_model.SelectionManager
    selection_data_part = selection_manager_part.CreateSelectData
    frame_path_list = sw_model.GetPathName.split('\\')
    assembly_name = frame_path_list[-1].split('.')[0]
    frame_path = '\\'.join(frame_path_list[:-1]) + '\\Детали\\' + ' '.join(assembly_name.split(' ')[0:2]) \
                 + f'\\{" ".join(assembly_name.split(" ")[2:])}'
    create_path_frame(frame_path)
    sw_model.Extension.SelectByID2('Твердые тела', 'BDYFOLDER', 0, 0, 0, False, 0, vt_dispatch, 0)
    bodies = selection_manager_part.GetSelectedObject6(1, -1)
    bodies.GetSpecificFeature2.SetAutomaticCutList(True)
    bodies.GetSpecificFeature2.SetAutomaticUpdate(True)
    sw_model.ClearSelection2(True)
    bodies = bodies.GetFirstSubFeature
    while True:
        if bodies is None:
            break
        bodies_count = bodies.GetSpecificFeature2.GetBodyCount
        if not bodies_count:
            bodies = bodies.GetNextSubFeature
            continue
        body = bodies.GetSpecificFeature2.GetBodies[0]
        body.Select2(True, selection_data_part)
        sw_model.SaveToFile3(f'{frame_path}\\{bodies.Name.replace("<", "(").replace(">", ")")}_{bodies_count}.SLDPRT',
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


def drill_sheet():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 3:
        sw_app.SendmsgToUser('Активен не чертеж')
        print('Активен не чертеж')
        return
    if sw_model.SelectionManager.GetSelectedObjectType3(1, -1) != 12:
        sw_app.SendmsgToUser('Не выбран вид чертежа')
        print('Не выбран вид чертежа')
        return
    vt_dispatch = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    arg1 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 2)
    arg2 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128)
    create_drill_sheet(sw_app, sw_model, vt_dispatch, arg1, arg2)


drill_sheet()
