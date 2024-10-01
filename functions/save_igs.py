import pythoncom

from functions.create_drill_sheet import get_ready
from functions.general_functions import create_app_model, create_com, check_assembly, clear_path


def check_path_igs(sw_app, components) -> tuple:
    """check and get directory"""

    component_path: str = check_path_pipe(sw_app, components)
    if not component_path:
        return ()

    path_list: list = component_path.split('\\')
    assembly_name: str = path_list[-2]
    path_list[-1] = 'Трубы IGS'
    path: str = '\\'.join(path_list)
    return assembly_name, path


def check_path_pipe(sw_app, components):
    """check pipe for standard"""

    count_standard = 0
    for component in components:
        if component.Name2.startswith('Труба'):
            component_path: str = component.GetPathName
            if 'Библиотека Solid Works НОВАЯ' in component_path:
                count_standard += 1
    if count_standard:
        sw_app.SendmsgToUser(f'⛔⛔ Труба из библиотеки {count_standard} шт. ⛔⛔')
        return
    return component_path


def get_count_tube(components: list) -> tuple[dict[str, dict[str, int]], dict[str, int]]:
    tubes: dict[str, dict[str, int]] = {}
    accounting: dict[str, int] = {}
    for component in components:
        if component.Name2.startswith('Труба'):
            name: str = component.Name2.split('-')[0]
            conf: str = component.ReferencedConfiguration
            if not tubes.get(name):
                tubes[name] = {conf: 1}
            else:
                tubes[name][conf] = tubes[name].setdefault(conf, 0) + 1
        elif component.Name2.startswith('УУТЭ'):
            name: str = f'{component.ReferencedConfiguration.split()[1]} {component.Name2.split()[1]}'
            accounting[name] = accounting.setdefault(name, 0) + 1
    return tubes, accounting


def create_pipe_igs(sw_app, path: str, pipes: dict[str, dict[str, int]], arg1, arg2):
    """open and save IGS pipe part"""

    path_pipe_list: list = path.split('\\')
    if get_ready():
        return
    path_tube: str = '\\'.join(path_pipe_list[:-1])
    for pipe, configurations in pipes.items():
        model = sw_app.OpenDoc6(f'{path_tube}\\{pipe}.SLDPRT', 1, 2, '', arg1, arg2)
        for configuration, count in configurations.items():
            if (model.ShowConfiguration2(configuration) == False and
                    model.ConfigurationManager.ActiveConfiguration.Name != configuration):
                return False
            pipe_new = pipe.replace('(Резьба зеркало)', '(З)').replace('(Плоскости от трубы)', '(Т)')
            model.SaveAs3(f'{path}\\{pipe_new} l={configuration} ({count} шт).igs', 0, 2)
        sw_app.CloseDoc(pipe)
    return True


def create_accounting_igs(sw_app, path: str, accounting: dict[str, int], arg1, arg2):
    """open and save IGS accounting part"""

    pipes_accounting: dict[str: list[list]] = {
        '25-20 Подающий': [['Труба ЭСВ Dn 65 (УУТЭ П)', '140 мм', 1], ['Труба ВГП Dn 20', '60 мм', 2]],
        '25-20 Обратный': [['Труба ЭСВ Dn 65 (УУТЭ О)', '140 мм', 1], ['Труба ВГП Dn 20', '60 мм', 2]],
        '32-20 Подающий': [['Труба ЭСВ Dn 65 (УУТЭ П)', '140 мм', 1], ['Труба ВГП Dn 20', '60 мм', 2]],
        '32-20 Обратный': [['Труба ЭСВ Dn 65 (УУТЭ О)', '140 мм', 1], ['Труба ВГП Dn 20', '60 мм', 2]],
        '40-20 Подающий': [['Труба ЭСВ Dn 65 (УУТЭ П)', '140 мм', 1], ['Труба ВГП Dn 20', '60 мм', 2]],
        '40-20 Обратный': [['Труба ЭСВ Dn 65 (УУТЭ О)', '140 мм', 1], ['Труба ВГП Dn 20', '60 мм', 2]],
        '40-25 Подающий': [['Труба ЭСВ Dn 65 (УУТЭ П)', '140 мм', 1], ['Труба ВГП Dn 25', '65 мм', 2]],
        '40-25 Обратный': [['Труба ЭСВ Dn 65 (УУТЭ О)', '140 мм', 1], ['Труба ВГП Dn 25', '65 мм', 2]],
        '40-32 Подающий': [['Труба ЭСВ Dn 65 (УУТЭ П)', '140 мм', 1], ['Труба ВГП Dn 32', '80 мм', 2]],
        '40-32 Обратный': [['Труба ЭСВ Dn 65 (УУТЭ О)', '140 мм', 1], ['Труба ВГП Dn 32', '80 мм', 2]],
        '50-20 Подающий': [['Труба ЭСВ Dn 65 (УУТЭ П)', '140 мм', 1], ['Труба ВГП Dn 20', '60 мм', 2]],
        '50-20 Обратный': [['Труба ЭСВ Dn 65 (УУТЭ О)', '140 мм', 1], ['Труба ВГП Dn 20', '60 мм', 2]],
        '50-25 Подающий': [['Труба ЭСВ Dn 65 (УУТЭ П)', '140 мм', 1], ['Труба ВГП Dn 25', '65 мм', 2]],
        '50-25 Обратный': [['Труба ЭСВ Dn 65 (УУТЭ О)', '140 мм', 1], ['Труба ВГП Dn 25', '65 мм', 2]],
        '50-32 Подающий': [['Труба ЭСВ Dn 65 (УУТЭ П)', '140 мм', 1], ['Труба ВГП Dn 32', '80 мм', 2]],
        '50-32 Обратный': [['Труба ЭСВ Dn 65 (УУТЭ О)', '140 мм', 1], ['Труба ВГП Dn 32', '80 мм', 2]],
        '50-40 Подающий': [['Труба ЭСВ Dn 65 (УУТЭ П)', '140 мм', 1], ['Труба ВГП Dn 40', '90 мм', 2]],
        '50-40 Обратный': [['Труба ЭСВ Dn 65 (УУТЭ О)', '140 мм', 1], ['Труба ВГП Dn 40', '90 мм', 2]],
        '65-32 Подающий': [['Труба ВГП Dn 32', '80 мм', 2]],
        '65-32 Обратный': [['Труба ВГП Dn 32', '80 мм', 2]],
        '65-40 Подающий': [['Труба ВГП Dn 40', '90 мм', 2]],
        '65-40 Обратный': [['Труба ВГП Dn 40', '90 мм', 2]],
        '65-50 Подающий': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '65-50 Обратный': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '65-50 (Расх. фланцевый) Подающий': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '65-50 (Расх. фланцевый) Обратный': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '80-40 Подающий': [['Труба ВГП Dn 40', '90 мм', 2]],
        '80-40 Обратный': [['Труба ВГП Dn 40', '90 мм', 2]],
        '80-50 Подающий': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '80-50 Обратный': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '80-50 (Расх. фланцевый) Подающий': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '80-50 (Расх. фланцевый) Обратный': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '80-65 Подающий': [['Труба ЭСВ Dn 65', '155 мм', 2]],
        '80-65 Обратный': [['Труба ЭСВ Dn 65', '155 мм', 2]],
        '100-50 Подающий': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '100-50 Обратный': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '100-50 (Расх. фланцевый) Подающий': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '100-50 (Расх. фланцевый) Обратный': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '100-65 Подающий': [['Труба ЭСВ Dn 65', '155 мм', 2]],
        '100-65 Обратный': [['Труба ЭСВ Dn 65', '155 мм', 2]],
        '100-80 Подающий': [['Труба ЭСВ Dn 80', '185 мм', 2]],
        '100-80 Обратный': [['Труба ЭСВ Dn 80', '185 мм', 2]],
        '125-65 Подающий': [['Труба ЭСВ Dn 65', '155 мм', 2]],
        '125-65 Обратный': [['Труба ЭСВ Dn 65', '155 мм', 2]],
        '125-80 Подающий': [['Труба ЭСВ Dn 80', '185 мм', 2]],
        '125-80 Обратный': [['Труба ЭСВ Dn 80', '185 мм', 2]],
        '150-80 Подающий': [['Труба ЭСВ Dn 80', '185 мм', 2]],
        '150-80 Обратный': [['Труба ЭСВ Dn 80', '185 мм', 2]],
        '150-100 Подающий': [['Труба ЭСВ Dn 100', '230 мм', 2]],
        '150-100 Обратный': [['Труба ЭСВ Dn 100', '230 мм', 2]],
    }
    path_accounting_part: str = '\\\\192.168.1.14\\SolidWorks\\Библиотека Solid Works НОВАЯ\\УУТЭ\\Детали вторичные'
    path_accounting_pipe: str = '\\\\192.168.1.14\\SolidWorks\\Библиотека Solid Works НОВАЯ\\Металл\\Трубы'
    # path_accounting_part: str = 'D:\\Solid Works\\Библиотека Solid Works НОВАЯ\\УУТЭ\\Детали вторичные'
    # path_accounting_pipe: str = 'D:\\Solid Works\\Библиотека Solid Works НОВАЯ\\Металл\\Трубы'

    pipes_accounting_current: dict[str, dict[str, int]] = {}
    for name, count in accounting.items():
        for pipe in pipes_accounting.get(name):
            if not pipes_accounting_current.get(pipe[0]):
                pipes_accounting_current[pipe[0]] = {pipe[1]: pipe[2] * count}
            else:
                pipes_accounting_current[pipe[0]][pipe[1]] = pipes_accounting_current[pipe[0]].setdefault(pipe[1], 0) + \
                                                             pipe[2] * count

    for pipe, configurations in pipes_accounting_current.items():
        if pipe.endswith(('(УУТЭ П)', '(УУТЭ О)')):
            path_model = f'{path_accounting_part}\\{pipe}.SLDPRT'
        else:
            path_model = f'{path_accounting_pipe}\\{pipe}.SLDPRT'
        model = sw_app.OpenDoc6(path_model, 1, 2, '', arg1, arg2)
        for configuration, count in configurations.items():
            if (model.ShowConfiguration2(configuration) == False and
                    model.ConfigurationManager.ActiveConfiguration.Name != configuration):
                return False
            pipe_new = pipe.replace('(Резьба зеркало)', '(З)').replace('(Плоскости от трубы)', '(Т)')
            model.SaveAs3(f'{path}\\{pipe_new}  l={configuration} ({count} шт).igs', 0, 2)
        sw_app.CloseDoc(pipe)
    return True


def main_save_igs():
    """initialization SW and main"""

    sw_app, sw_model = create_app_model()

    if not check_assembly(sw_app, sw_model):
        return

    components: list = sw_model.GetComponents(True)

    assembly_name, path = check_path_igs(sw_app, components)
    if not assembly_name:
        return

    tubes, accounting = get_count_tube(components)

    path_assembly: str = sw_model.GetPathName
    sw_app.CloseDoc(assembly_name)

    clear_path(path)
    arg1 = create_com(2, pythoncom.VT_BYREF, pythoncom.VT_I4)
    arg2 = create_com(128, pythoncom.VT_BYREF, pythoncom.VT_I4)

    if not create_pipe_igs(sw_app, path, tubes, arg1, arg2):
        sw_app.SendmsgToUser('⛔⛔ Ошибка в конфигурациях, будут расхождения длин ⛔⛔')
        return

    if not create_accounting_igs(sw_app, path, accounting, arg1, arg2):
        sw_app.SendmsgToUser('⛔⛔ Ошибка в конфигурациях, будут расхождения длин ⛔⛔')
        return

    sw_app.OpenDoc6(path_assembly, 2, 32, '', arg1, arg2)

    sw_app.SendmsgToUser('IGS успешно сохранены')
