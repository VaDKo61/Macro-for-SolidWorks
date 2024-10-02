import os
import win32com.client


def create_app_model():
    """create sw_app and sw_model"""

    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    sw_model = sw_app.ActiveDoc
    return sw_app, sw_model


def check_assembly(sw_app, sw_model) -> bool:
    """check object for assembly"""

    if sw_model.GetType != 2:
        sw_app.SendmsgToUser('⛔⛔ Активна не сборка ⛔⛔')
        return False
    return True


def check_unselect_element(sw_app, sw_model) -> bool:
    """check selected elements or not"""

    if not sw_model.SelectionManager.GetSelectedObjectCount2(-1):
        sw_app.SendmsgToUser('⛔⛔ Не выбраны трубы ⛔⛔')
        return False
    return True


def check_edge(sw_app, sel_manager, index) -> bool:
    """check object for edge"""

    if sel_manager.GetSelectedObjectType3(index, -1) != 1:
        sw_app.SendmsgToUser('⛔⛔ Не выбрана кромка врезаемой трубы ⛔⛔')
        return False
    return True


def check_surface(sw_app, sel_manager, index) -> bool:
    """check object for surface"""

    if sel_manager.GetSelectedObjectType3(index, -1) != 2:
        sw_app.SendmsgToUser('⛔⛔ Не выбрана поверхность трубы ⛔⛔')
        return False
    return True


def create_check_path(path: str):
    """check and create path"""

    try:
        os.makedirs(path)
    except FileExistsError:
        return False


def create_com(value, *args):
    """create COM object"""

    if len(args) == 2:
        return win32com.client.VARIANT(args[0] | args[1], value)


def clear_path(path: str):
    if os.path.isdir(path):
        for file in os.listdir(path):
            os.remove(f'{path}\\{file}')


def create_select_man_data(sw_model):
    """create selection manager and data"""

    sel_manager = sw_model.SelectionManager
    sel_data = sel_manager.CreateSelectData
    return sel_manager, sel_data


def add_unique_conf(sw_model_pipe, name_conf, all_name_conf) -> str:
    """add in pipe unique configurations"""

    for i in range(2, 51):
        name_new_conf: str = f'{name_conf}({i})'
        if name_new_conf not in all_name_conf:
            sw_model_pipe.ConfigurationManager.AddConfiguration2(name_new_conf, '', '', 128,
                                                                 name_conf, '', True)
            return name_new_conf
