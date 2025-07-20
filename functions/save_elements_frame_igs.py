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
        clear_path(path)
        sw_app.SendmsgToUser('⛔⛔ Не удалось сохранить файлы ⛔⛔')
        return

    sw_app.SendmsgToUser('Элементы рамы успешно сохранены')
