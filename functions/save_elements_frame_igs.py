import pythoncom

from functions.archive.create_drill_sheet import get_ready
from functions.general_functions import create_check_path, clear_path, create_com, create_select_man_data, \
    create_app_model, check_part


def create_path_frame(sw_app, sw_model) -> str:
    """check and crete path for frame"""

    path_list: list = sw_model.GetPathName.split('\\')
    part_name: str = path_list.pop().split('.')[0]
    part_name_list: list = part_name.split()
    if part_name.count('.') > 0:
        sw_app.SendmsgToUser('⛔⛔ Название содержит точки ⛔⛔')
        return ''
    if part_name.count(' ') != 2:
        sw_app.SendmsgToUser('⛔⛔ Название не по шаблону ⛔⛔')
        return ''
    if '№' not in part_name:
        sw_app.SendmsgToUser('⛔⛔ Название не по шаблону ⛔⛔')
        return ''
    path_list.extend(('Лазер', 'Трубы', f'{part_name_list[0]} Лазер', f'{part_name_list[1]} {part_name_list[2]} IGS'))
    path: str = '\\'.join(path_list)
    if not create_check_path(path):
        clear_path(path)
    return path


def save_elements_igs(sw_app, sw_model, path) -> bool:
    """save elements frame in igs"""

    arg1 = create_com(None, pythoncom.VT_DISPATCH)
    sw_model.ClearSelection2(True)
    sel_manager, sel_data = create_select_man_data(sw_model)
    sw_model.Extension.SelectByID2('Твердые тела', 'BDYFOLDER', 0, 0, 0, False, 0, arg1, 0)
    bodies = sel_manager.GetSelectedObject6(1, -1)
    bodies.GetSpecificFeature2.SetAutomaticCutList(True)
    bodies.GetSpecificFeature2.SetAutomaticUpdate(True)
    sw_model.ClearSelection2(True)
    if get_ready():
        return True
    bodies = bodies.GetFirstSubFeature
    arg2 = create_com(2, pythoncom.VT_BYREF, pythoncom.VT_I4)
    arg3 = create_com(128, pythoncom.VT_BYREF, pythoncom.VT_I4)
    arg4 = create_com(None, pythoncom.VT_BYREF, pythoncom.VT_BSTR)
    arg5 = create_com(None, pythoncom.VT_BYREF, pythoncom.VT_BSTR)
    arg6 = create_com(True, pythoncom.VT_BYREF, pythoncom.VT_BOOL)
    arg7 = create_com(True, pythoncom.VT_BYREF, pythoncom.VT_BOOL)
    while True:
        if bodies is None:
            return True
        bodies_count = bodies.GetSpecificFeature2.GetBodyCount
        if not bodies_count:
            bodies = bodies.GetNextSubFeature
            continue
        body = bodies.GetSpecificFeature2.GetBodies[0]
        bodies.CustomPropertyManager.Get6('Длина', False, arg4, arg5, arg6, arg7)
        body.Select2(False, sel_data)
        name_element: str = bodies.Name.replace("<", "(").replace(">", ")")
        path_element: str = '{}\\{} l={} мм ({} шт.).IGS'.format(path, name_element, arg5.value, bodies_count)
        sw_model.SaveToFile3(path_element, 2, 2, False, False, arg2, arg3)
        sw_app.CloseDoc('')
        sw_model.ClearSelection2(True)
        bodies = bodies.GetNextSubFeature


def main_save_frame_igs():
    """initialization SW and main"""

    sw_app, sw_model = create_app_model()

    if not check_part(sw_app, sw_model):
        return

    path = create_path_frame(sw_app, sw_model)
    if not path:
        sw_app.SendmsgToUser('⛔⛔ Не удалось сформировать путь ⛔⛔')
        return

    if not save_elements_igs(sw_app, sw_model, path):
        sw_app.SendmsgToUser('⛔⛔ Не удалось сохранить файлы ⛔⛔')
        return

    sw_app.SendmsgToUser('Элементы рамы успешно сохранены')
