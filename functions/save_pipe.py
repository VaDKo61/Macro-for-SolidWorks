import pythoncom

from functions.archive.create_drill_sheet import get_ready
from functions.general_functions import check_assembly, create_check_path, create_app_model, create_com, save_assembly


def create_path(sw_app, sw_model) -> str:
    """check and crete path"""

    path_list: list = sw_model.GetPathName.split('\\')
    if path_list[-1].count('.') > 1:
        sw_app.SendmsgToUser('⛔⛔ Название содержит точки ⛔⛔')
        return ''
    assembly_name: str = path_list.pop().split('.')[0]
    if assembly_name.count(' ') != 1:
        sw_app.SendmsgToUser('⛔⛔ Название не по шаблону ⛔⛔')
        return ''
    elif not assembly_name.endswith('Лазер'):
        sw_app.SendmsgToUser('⛔⛔ В названии нет: "Лазер" ⛔⛔')
        return ''

    path_list.append('Трубы')
    path_list.append(assembly_name)
    path: str = '\\'.join(path_list)

    create_check_path(f'{path}\\Трубы IGS')
    return path


def save_pipe(sw_model, path: str):
    """save pipe in directory assembly"""

    pipes: list = []
    if get_ready():
        return
    for component in sw_model.GetComponents(True):
        component_name: str = component.name2
        if component_name.startswith('Труба'):
            component_name = component_name.split('-')[0]
            if component_name not in pipes:
                pipes.append(component_name)
                pipe = component.GetModelDoc2
                pipe.SaveAs3(f'{path}\\{component_name}.SLDPRT', 0, 8)


def main_save_pipe():
    """initialization SW and main"""

    sw_app, sw_model = create_app_model()

    arg1 = create_com(2, pythoncom.VT_BYREF, pythoncom.VT_I4)
    arg2 = create_com(128, pythoncom.VT_BYREF, pythoncom.VT_I4)

    if not check_assembly(sw_app, sw_model):
        return

    path: str = create_path(sw_app, sw_model)
    if not path:
        return

    save_pipe(sw_model, path)

    if not save_assembly(sw_model, arg1, arg2):
        sw_app.SendmsgToUser('⛔⛔ Сборка не сохранилась ⛔⛔')
        return

    sw_app.SendmsgToUser('Трубы успешно сохранены')
