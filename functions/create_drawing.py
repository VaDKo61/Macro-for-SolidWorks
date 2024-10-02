import pythoncom
import win32com.client

from functions.archive.create_drill_sheet import get_ready


def create_drawing(sw_app, sw_model, vt_dispatch):
    """create 2-5 lists"""

    assembly_path = sw_model.GetCurrentSheet.GetViews[-1].GetReferencedModelName
    view_scale = sw_model.GetCurrentSheet.GetViews[-1].ScaleDecimal
    first_sheet = sw_model.GetCurrentSheet.GetName
    sw_model.Extension.SelectByID2(sw_model.GetCurrentSheet.GetViews[-1].GetName2, 'DRAWINGVIEW', 0, 0, 0, False, 0,
                                   vt_dispatch, 0)
    sw_model.EditDelete()
    sheet_names: list = ['Изом2', 'Изом3', 'Изом4', 'Габариты']
    if get_ready():
        return
    for name in sheet_names:
        current_name = sw_model.GetCurrentSheet.GetName
        sw_model.Extension.SelectByID2(current_name, 'SHEET', 0, 0, 0, False, 0, vt_dispatch, 0)
        sw_model.EditCopy()
        sw_model.PasteSheet(1, 1)
        sw_model.GetCurrentSheet.SetName(name)
    sheet_names.insert(0, first_sheet)

    # add draw view
    view_names: list = ['*Изометрия', 'Изом 2', 'Изом 3', 'Изом 4']
    for i in zip(sheet_names[:4], view_names[:4]):
        sw_model.ActivateSheet(i[0])
        current_view = sw_model.CreateDrawViewFromModelView3(assembly_path, i[1], 0.21, 0.1485, 0)
        current_view.SetDisplayMode4(False, 3, False, True, False)
        current_view.ScaleDecimal = view_scale
    sw_model.ActivateSheet(sheet_names[-1])
    sw_model.Create1stAngleViews2(assembly_path)
    sw_app.SendmsgToUser('Листы успешно добавлены')


def drawing():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    vt_dispatch = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 3:
        sw_app.SendmsgToUser('Активен не чертеж')
        return
    create_drawing(sw_app, sw_model, vt_dispatch)
