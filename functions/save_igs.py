import os

import pythoncom
import win32com.client


def get_path_igs(assembly_path: str) -> tuple:
    """Create or clear and get directory"""
    path_list: list = assembly_path.split('\\')
    assembly_name: str = path_list.pop().split('.')[0]
    path_list.append('Трубы')
    path_list.append(assembly_name)
    path_list.append('IGS')
    path: str = '\\'.join(path_list)
    if os.path.isdir(path):
        for file in os.listdir(path):
            os.remove(f'{path}\\{file}')
        else:
            print('Директория была очищена от IGS файлов')
    else:
        os.makedirs(path)
        print(f'Директория {path} была создана')
    return assembly_name, path


def get_count_tube(components: list) -> dict[str, dict[str, int]]:
    tubes: dict[str, dict[str, int]] = {}
    for component in components:
        if component.Name2.startswith('Труба') or component.Name2.startswith('Ниппель'):
            name: str = component.Name2.split('-')[0]
            conf: str = component.ReferencedConfiguration
            if not tubes.get(name):
                tubes[name] = {conf: 1}
            else:
                tubes[name][conf] = tubes[name].setdefault(conf, 0) + 1
    return tubes


def create_igs(sw_app, assembly_name: str, path: str, tubes: dict[str, dict[str, int]], arg5, arg6):
    """Create IGS, open tube part"""
    sw_app.CloseDoc(assembly_name)
    path_tube_list: list = path.split('\\')
    path_tube: str = '\\'.join(path_tube_list[:-1])
    path_assembly = '\\'.join(path_tube_list[:-3])
    for tube, configurations in tubes.items():
        model = sw_app.OpenDoc6(f'{path_tube}\\{tube}.SLDPRT', 1, 2, '', arg5, arg6)
        for configuration, count in configurations.items():
            model.ShowConfiguration2(configuration)
            thread_1 = model.FeatureByName('Бобышка-Вытянуть2')
            if thread_1:
                thread_1.SetSuppression2(0, 1)
            thread_1 = model.FeatureByName('Бобышка-Вытянуть3')
            if thread_1:
                thread_1.SetSuppression2(0, 1)
            tube_new = tube.replace('(Резьба зеркало)', '(З)').replace('(Плоскости от трубы)', '(Т)')
            model.SaveAs3(f'{path}\\{tube_new} l={configuration} ({count} шт).igs', 0, 2)
        sw_app.CloseDoc(tube)
    sw_app.OpenDoc6(f'{path_assembly}\\{assembly_name}.SLDASM', 2, 32, '', arg5, arg6)


def save_igs():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    arg1 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 2)
    arg2 = win32com.client.VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, 128)
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 2:
        sw_app.SendmsgToUser('Активна не сборка')
        print('Активна не сборка')
        return
    assembly_name, path = get_path_igs(sw_model.GetPathName)
    components = sw_model.GetComponents(True)

    tubes: dict[str, dict[str, int]] = get_count_tube(components)

    create_igs(sw_app, assembly_name, path, tubes, arg1, arg2)

    sw_app.SendmsgToUser('IGS успешно сохранены')
    print('IGS успешно сохранены')
