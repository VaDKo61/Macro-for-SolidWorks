import pythoncom

from functions.archive.create_drill_sheet import get_ready
from functions.general_functions import create_app_model, check_drawing, create_com


def get_scale(sw_model) -> float:
    """get scale active sheet"""

    return sw_model.GetCurrentSheet.GetViews[-1].ScaleDecimal


def get_name_path(sw_model) -> tuple:
    """get assembly path active view and name sheet"""

    return sw_model.GetCurrentSheet.GetViews[-1].GetReferencedModelName, sw_model.GetCurrentSheet.GetName


def delete_view(sw_model, arg1):
    """delete view"""

    name_view: str = sw_model.GetCurrentSheet.GetViews[-1].GetName2
    sw_model.Extension.SelectByID2(name_view, 'DRAWINGVIEW', 0, 0, 0, False, 0, arg1, 0)
    sw_model.EditDelete()
    sw_model.ClearSelection2(True)
    return


def add_sheet(sw_model, arg1, name_sheet) -> list:
    """add sheet and name hire"""

    sheet_names: list = ['Изом2', 'Изом3', 'Изом4', 'Габариты']
    if get_ready():
        return True
    for name in sheet_names:
        current_name = sw_model.GetCurrentSheet.GetName
        sw_model.Extension.SelectByID2(current_name, 'SHEET', 0, 0, 0, False, 0, arg1, 0)
        sw_model.EditCopy()
        sw_model.PasteSheet(1, 1)
        sw_model.GetCurrentSheet.SetName(name)
    sheet_names.insert(0, name_sheet)
    return sheet_names


def add_view(sw_model, sheet_names: list, assembly_path: str, scale: float):
    """add view in sheets"""

    view_names: list = ['*Изометрия', 'Изом 2', 'Изом 3', 'Изом 4']
    for i in zip(sheet_names[:4], view_names[:4]):
        sw_model.ActivateSheet(i[0])
        current_view = sw_model.CreateDrawViewFromModelView3(assembly_path, i[1], 0.21, 0.1485, 0)
        current_view.SetDisplayMode4(False, 3, False, True, False)
        current_view.ScaleDecimal = scale
    sw_model.ActivateSheet(sheet_names[-1])
    sw_model.Create1stAngleViews2(assembly_path)
    return


def main_create_drawing():
    """initialization SW and main"""

    sw_app, sw_model = create_app_model()

    if not check_drawing(sw_app, sw_model):
        return

    scale = get_scale(sw_model)
    if not scale:
        sw_app.SendmsgToUser('⛔⛔ На странице не вставлены виды ⛔⛔')
        return

    assembly_path, name_sheet = get_name_path(sw_model)

    arg1 = create_com(None, pythoncom.VT_DISPATCH)
    delete_view(sw_model, arg1)

    sheet_names = add_sheet(sw_model, arg1, name_sheet)

    add_view(sw_model, sheet_names, assembly_path, scale)

    sw_app.SendmsgToUser('Листы успешно добавлены')
