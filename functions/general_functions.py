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