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
    bodies = bodies.GetFirstSubFeature
    # print(getattr(pythoncom, "VT_BSTR"))
    # # arg5 = win32com.client.VARIANT(pythoncom.VT_BSTR, "")
    # # print(bodies.GetFirstSubFeature.CustomPropertyManager.Get6('ДЛИНА', True, arg5, arg5, True, True))
    # a = ''
    # arg3 = win32com.client.VARIANT(pythoncom.VT_R8 | pythoncom.VT_BSTR, '')
    # print(bodies.GetFirstSubFeature.CustomPropertyManager.Get5("Длина", False, arg3, 0, 0))
    # print(bodies.GetFirstSubFeature.CustomPropertyManager.Get('TOTAL LENGTH'))
    # print(bodies.GetFirstSubFeature.Parameter('D1'))

    sw_model.ClearSelection2(True)
    # while not bodies.GetNextFeature.Name.startswith('Отрезок'):
    #     bodies = bodies.GetNextFeature
    while True:
        if bodies is None:
            break

        # args = [VT_EMPTY, 'VT_NULL', 'VT_I2', 'VT_I4', 'VT_R4', 'VT_R8', 'VT_CY', 'VT_DATE', 'VT_BSTR', 'VT_DISPATCH',
        #         'VT_ERROR', 'VT_BOOL', 'VT_VARIANT', 'VT_UNKNOWN', 'VT_DECIMAL', 'VT_I1', 'VT_UI1', 'VT_UI2', 'VT_UI4',
        #         'VT_I8', 'VT_UI8', 'VT_INT', 'VT_UINT', 'VT_VOID', 'VT_HRESULT', 'VT_PTR',
        #         'VT_SAFEARRAY', 'VT_CARRAY', 'VT_USERDEFINED', 'VT_LPSTR', 'VT_LPWSTR', 'VT_RECORD', 'VT_INT_PTR',
        #         'VT_UINT_PTR', 'VT_ARRAY']
        # arg3 = win32com.client.VARIANT(pythoncom.VT_BOOL | pythoncom.VT_BSTR, '')
        # print(bodies.CustomPropertyManager.Get5("Длина", False, arg3, 0, 0))

        bodies_count = bodies.GetSpecificFeature2.GetBodyCount
        if not bodies_count:
            bodies = bodies.GetNextSubFeature
            continue
        body = bodies.GetSpecificFeature2.GetBodies[0]
        body.Select2(True, selection_data)
        sw_model.SaveToFile3(f'{frame_path}\\{bodies.Name.replace("<", "(").replace(">", ")")} ({bodies_count} шт).IGS',
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

