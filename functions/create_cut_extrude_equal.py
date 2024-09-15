import win32com.client

from functions.create_drill_sheet import get_ready


def create_cut_extrude_equal(sw_app, sw_model):
    """Create sketch and cut extrude equal tube"""
    selection_manager = sw_model.SelectionManager
    selection_data = selection_manager.CreateSelectData
    edges = selection_manager.GetSelectedObject6(1, -1)
    component_edges = selection_manager.GetSelectedObjectsComponent4(1, -1)
    length_comp_edges: float = int(component_edges.ReferencedConfiguration.split()[0]) / 1000 + 0.01
    selection_manager.DeSelect2(1, -1)
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
            print('Не удалось добавить конфигурацию')
            return
        sw_model.EditAssembly()
        cut_in_pipe.ReferencedConfiguration = name_new_configurations
        cut_in_pipe.Select4(True, selection_data, False)
        sw_model.EditPart()

    # create plane
    edges.Select4(True, selection_data)
    plane = sw_model.FeatureManager.InsertRefPlane(4, 0, 0, 0, 0, 0)

    # create sketch
    plane.Select4(True, selection_data)
    if get_ready():
        return
    sw_model.SketchManager.InsertSketch(True)
    edges.Select4(True, selection_data)
    sw_model.SketchManager.SketchUseEdge3(False, False)

    # create cut extrude
    feature_cut = sw_model.FeatureManager.FeatureCut4(True, False, False, 0, 0, length_comp_edges, 0.001,
                                                      False, False, False, False, 0, 0, True, False, False, False,
                                                      False, True, True, True, True, False, 0, 0, False, False)
    if not feature_cut:
        sw_model.FeatureManager.FeatureCut4(True, False, False, 0, 0, length_comp_edges, 0.001, False, False,
                                            False, False, 0, 0, False, False, False, False, False, True, True, True,
                                            True, False, 0, 0, False, False)

    sw_model.ClearSelection2(True)
    sw_model.EditAssembly()
    a = sw_model.EditRebuild3
    sw_app.SendmsgToUser('Отверстие успешно создано')
    print('Отверстие успешно создано')


def cut_extrude_equal():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 2:
        sw_app.SendmsgToUser('Активна не сборка')
        print('Активна не сборка')
        return
    if sw_model.SelectionManager.GetSelectedObjectType3(1, -1) != 1:
        sw_app.SendmsgToUser('Не выбрана кромка врезаемой трубы')
        print('Не выбрана кромка врезаемой трубы')
        return
    if sw_model.SelectionManager.GetSelectedObjectType3(2, -1) != 2:
        sw_app.SendmsgToUser('Не выбрана поверхность трубы для отверстия')
        print('Не выбрана поверхность трубы для отверстия')
        return
    create_cut_extrude_equal(sw_app, sw_model)
