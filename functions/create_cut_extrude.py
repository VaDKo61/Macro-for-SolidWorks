import win32com.client

from functions.create_drill_sheet import get_ready


def create_any_cut_extrude(sw_app, sw_model, count_select, kip):
    """Create sketch and any cut extrude"""
    selection_manager = sw_model.SelectionManager
    selection_data = selection_manager.CreateSelectData
    edges: list = []
    for i in range(1, count_select):
        edges.append(selection_manager.GetSelectedObject6(1, -1))
        selection_manager.DeSelect2(1, -1)
    surface = selection_manager.GetSelectedObject6(1, -1)
    try:
        selection_data_sur = selection_manager.CreateSelectData
        surface.Select4(True, selection_data_sur)
    except BaseException:
        sw_app.SendmsgToUser('Ошибка, запустите заново')
        return

    cut_in_pipe = selection_manager.GetSelectedObjectsComponent4(1, -1)
    sw_model.EditPart()
    sw_model_part = sw_model.GetEditTarget

    # create configuration
    name_active_configuration: str = sw_model_part.ConfigurationManager.ActiveConfiguration.Name
    if not name_active_configuration.endswith(')'):
        name_configurations: str = sw_model_part.GetConfigurationNames
        for i in range(2, 51):
            name_new_configurations: str = f'{name_active_configuration}({i})'
            if name_new_configurations not in name_configurations:
                sw_model_part.ConfigurationManager.AddConfiguration2(name_new_configurations, '', '', 128,
                                                                     name_active_configuration, '', True)
                break
        else:
            sw_app.SendmsgToUser('Не удалось добавить конфигурацию')
            return
        sw_model.EditAssembly()
        cut_in_pipe.ReferencedConfiguration = name_new_configurations
        cut_in_pipe.Select4(True, selection_data, False)
        sw_model.EditPart()

    # create plane
    edges[0].Select4(True, selection_data)
    if get_ready():
        return
    plane = sw_model.FeatureManager.InsertRefPlane(4, 0, 0, 0, 0, 0)

    # create sketch
    plane.Select4(True, selection_data)
    sw_model.SketchManager.InsertSketch(True)
    for edge in edges:
        edge.Select4(True, selection_data)
    sw_model.SketchManager.SketchUseEdge3(False, False)

    if kip:
        edges_uses: tuple = sw_model.GetActiveSketch2.GetSketchSegments
        for edges_use in edges_uses:
            edges_use.ConstructionGeometry = True
            edges_use.Select4(True, selection_data)
            sw_model.SketchManager.SketchOffset2(0.001, False, True, 0, 0, True)

    # create cut extrude
    surface.Select4(True, selection_data)
    feature_cut = sw_model.FeatureManager.FeatureCut4(True, False, False, 5, 0, 0.005, 0.001, False, False, False,
                                                      False, 0, 0, True, False, False, False, False, True, True, True,
                                                      True, False, 0, 0, False, False)
    if not feature_cut:
        sw_model.FeatureManager.FeatureCut4(True, False, False, 5, 0, 0.005, 0.001, False, False, False,
                                            False, 0, 0, False, False, False, False, False, True, True, True,
                                            True, False, 0, 0, False, False)
    sw_model.ClearSelection2(True)
    sw_model.EditAssembly()
    a = sw_model.EditRebuild3
    sw_app.SendmsgToUser('Отверстие успешно создано')


def any_cut_extrude():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 2:
        sw_app.SendmsgToUser('Активна не сборка')
        return
    if sw_model.SelectionManager.GetSelectedObjectType3(1, -1) != 1:
        sw_app.SendmsgToUser('Не выбрана кромка врезаемой трубы')
        return
    count_select = sw_model.SelectionManager.GetSelectedObjectCount2(-1)
    if sw_model.SelectionManager.GetSelectedObjectType3(count_select, -1) != 2:
        sw_app.SendmsgToUser('Не выбрана поверхность трубы для отверстия')
        return
    create_any_cut_extrude(sw_app, sw_model, count_select, kip=False)


def any_cut_extrude_kip():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 2:
        sw_app.SendmsgToUser('Активна не сборка')
        return
    if sw_model.SelectionManager.GetSelectedObjectType3(1, -1) != 1:
        sw_app.SendmsgToUser('Не выбрана кромка врезаемой трубы')
        return
    count_select = sw_model.SelectionManager.GetSelectedObjectCount2(-1)
    if sw_model.SelectionManager.GetSelectedObjectType3(count_select, -1) != 2:
        sw_app.SendmsgToUser('Не выбрана поверхность трубы для отверстия')
        return
    create_any_cut_extrude(sw_app, sw_model, count_select, kip=True)