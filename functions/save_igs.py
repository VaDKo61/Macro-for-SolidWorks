import os

import pythoncom
from future.backports import count

from functions.create_drill_sheet import get_ready
from functions.general_functions import create_app_model, create_com, check_assembly


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
        if component.Name2.startswith('Труб'):
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


def create_igs(sw_app, assembly_name: str, path: str, tubes: dict[str, dict[str, int]], accounting: dict[str, int],
               arg5, arg6):
    """Create IGS, open tube part"""
    sw_app.CloseDoc(assembly_name)
    path_tube_list: list = path.split('\\')
    if get_ready():
        return
    path_tube: str = '\\'.join(path_tube_list[:-1])
    path_assembly = '\\'.join(path_tube_list[:-3])

    # save igs tube
    for tube, configurations in tubes.items():
        model = sw_app.OpenDoc6(f'{path_tube}\\{tube}.SLDPRT', 1, 2, '', arg5, arg6)
        for configuration, count in configurations.items():
            model.ShowConfiguration2(configuration)
            # thread_1 = model.FeatureByName('Бобышка-Вытянуть2')
            # if thread_1:
            #     thread_1.SetSuppression2(0, 1)
            # thread_1 = model.FeatureByName('Бобышка-Вытянуть3')
            # if thread_1:
            #     thread_1.SetSuppression2(0, 1)
            tube_new = tube.replace('(Резьба зеркало)', '(З)').replace('(Плоскости от трубы)', '(Т)')
            model.SaveAs3(f'{path}\\{tube_new} l={configuration} ({count} шт).igs', 0, 2)
        sw_app.CloseDoc(tube)

    # save igs accounting
    tubes_accounting: dict[str: list[list]] = {
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
        '65-32': [['Труба ВГП Dn 32', '80 мм', 2]],
        '65-40': [['Труба ВГП Dn 40', '90 мм', 2]],
        '65-50': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '65-50 (Расх. фланцевый)': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '80-40': [['Труба ВГП Dn 40', '90 мм', 2]],
        '80-50': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '80-50 (Расх. фланцевый)': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '80-65': [['Труба ЭСВ Dn 65', '155 мм', 2]],
        '100-50': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '100-50 (Расх. фланцевый)': [['Труба ЭСВ Dn 50', '125 мм', 2]],
        '100-65': [['Труба ЭСВ Dn 65', '155 мм', 2]],
        '100-80': [['Труба ЭСВ Dn 80', '185 мм', 2]],
        '125-65': [['Труба ЭСВ Dn 65', '155 мм', 2]],
        '125-80': [['Труба ЭСВ Dn 80', '185 мм', 2]],
        '150-80': [['Труба ЭСВ Dn 80', '185 мм', 2]],
        '150-100': [['Труба ЭСВ Dn 100', '230 мм', 2]],
    }
    for name, count in accounting.items():
        for tube in tubes_accounting.get(name):
            if tube[0].endswith('(УУТЭ П)') or tube[0].endswith('(УУТЭ О)'):
                path_model = f'\\\\192.168.1.14\\SolidWorks\\Библиотека Solid Works НОВАЯ\\УУТЭ\\Детали вторичные' \
                             f'\\{tube[0]}.SLDPRT'
            else:
                path_model = f'\\\\192.168.1.14\\SolidWorks\\Библиотека Solid Works НОВАЯ\\Металл\\Трубы' \
                             f'\\{tube[0]}.SLDPRT'
            model = sw_app.OpenDoc6(path_model, 1, 2, '', arg5, arg6)
            model.ShowConfiguration2(tube[1])
            model.SaveAs3(f'{path}\\{tube[0]} l={tube[1]} ({count * tube[2]} шт).igs', 0, 2)
            sw_app.CloseDoc(tube[0])
    sw_app.OpenDoc6(f'{path_assembly}\\{assembly_name}.SLDASM', 2, 32, '', arg5, arg6)


def main_save_igs():
    """initialization SW and main"""

    sw_app, sw_model = create_app_model()

    arg1 = create_com(2, pythoncom.VT_BYREF, pythoncom.VT_I4)
    arg2 = create_com(128, pythoncom.VT_BYREF, pythoncom.VT_I4)

    if not check_assembly(sw_app, sw_model):
        return

    components: list = sw_model.GetComponents(True)

    assembly_name, path = check_path_igs(sw_app, components)
    if not assembly_name:
        return

    tubes: dict[str, dict[str, int]] = get_count_tube(components)[0]
    accounting: dict[str, int] = get_count_tube(components)[1]

    create_igs(sw_app, assembly_name, path, tubes, accounting, arg1, arg2)

    sw_app.SendmsgToUser('IGS успешно сохранены')


main_save_igs()