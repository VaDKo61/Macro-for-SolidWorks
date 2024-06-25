import pythoncom
import win32com.client


def create_cut_extrude(sw_app, sw_model):
    """Create sketch and cut extrude"""
    selection_manager = sw_model.SelectionManager
    edges = selection_manager.GetSelectedObject6(1, -1)
    selection_manager.DeSelect2(1, -1)
    cut_in_pipe = selection_manager.GetSelectedObjectsComponent4(1, -1)
    sw_model.EditPart()
    sw_model_part = sw_model.GetEditTarget

    # create configuration
    name_active_configuration: str = sw_model_part.ConfigurationManager.ActiveConfiguration.Name
    name_configurations: str = sw_model_part.GetConfigurationNames
    for i in range(2, 51):
        name_new_configurations: str = f'{name_active_configuration}({i})'
        if name_new_configurations not in name_configurations:
            sw_model_part.ConfigurationManager.AddConfiguration2(name_new_configurations, '', '', 128,
                                                                 name_active_configuration, '', True)
            break
    else:
        sw_app.SendmsgToUser('Не удалось добавить конфигурацию')
        print('Не удалось добавить конфигурацию')
        return
    sw_model.EditAssembly()
    cut_in_pipe.ReferencedConfiguration = name_new_configurations
    selection_data = selection_manager.CreateSelectData
    cut_in_pipe.Select4(True, selection_data, False)
    sw_model.EditPart()

    # create sketch
    sw_model.SketchManager.Insert3DSketch(True)
    edges.Select4(True, selection_data)
    sw_model.SketchManager.ConstructionGeometry = True


def cut_extrude():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    vt_dispatch = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 2:
        sw_app.SendmsgToUser('Активна не сборка')
        print('Активна не сборка')
        return
    print(sw_model.SelectionManager.GetSelectedObjectType3(1, -1))
    if sw_model.SelectionManager.GetSelectedObjectType3(1, -1) != 1:
        sw_app.SendmsgToUser('Не выбрана кромка врезаемой трубы')
        print('Не выбрана кромка врезаемой трубы')
        return
    if sw_model.SelectionManager.GetSelectedObjectType3(2, -1) != 2:
        sw_app.SendmsgToUser('Не выбрана поверхность трубы для отверстия')
        print('Не выбрана поверхность трубы для отверстия')
        return
    create_cut_extrude(sw_app, sw_model)


cut_extrude()
