import pythoncom
import win32com.client


def create_drawing(sw_app, sw_model, vt_dispatch):
    assembly_path: str = sw_model.GetPathName
    assembly_path_list: list = assembly_path.split('\\')[2:]
    engineer: str = assembly_path_list[2]
    template_path = f'\\\\{assembly_path_list[0]}\\{assembly_path_list[1]}\\{engineer}\\Шаблоны\\Чертеж сборки.DRWDOT'
    sw_draw = sw_app.NewDocument(template_path, 12, 0.42, 0.297)

    # create sheet
    sheet_names: list = ['*Изометрия', 'Изом 2', 'Изом 3', 'Изом 4', 'Габариты']
    for name in sheet_names[:4]:
        sw_draw.GetCurrentSheet.SetName(name)
        sw_draw.Extension.SelectByID2(name, 'SHEET', 0, 0, 0, False, 0, vt_dispatch, 0)
        sw_draw.EditCopy()
        sw_draw.PasteSheet(1, 1)
    sw_draw.GetCurrentSheet.SetName(sheet_names[4])

    # add draw view
    for name in sheet_names[:4]:
        sw_draw.ActivateSheet(name)
        current_view = sw_draw.CreateDrawViewFromModelView3(assembly_path, name, 0.21, 0.1485, 0)
        current_view.SetDisplayMode4(False, 3, True, True, False)
    sw_draw.ActivateSheet(sheet_names[4])
    sw_draw.Create1stAngleViews2(assembly_path)


def drawing():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    vt_dispatch = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 2:
        sw_app.SendmsgToUser('Активна не сборка')
        print('Активна не сборка')
        return
    create_drawing(sw_app, sw_model, vt_dispatch)


drawing()
