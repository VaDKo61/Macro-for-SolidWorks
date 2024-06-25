import win32com.client


def create_cut_extrude(sw_app, sw_model, kip):
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

    # create plane
    edges.Select4(True, selection_data)
    plane = sw_model.FeatureManager.InsertRefPlane(4, 0, 0, 0, 0, 0)

    # create sketch
    plane.Select4(True, selection_data)
    sw_model.SketchManager.InsertSketch(True)
    edges.Select4(True, selection_data)
    sw_model.SketchManager.SketchUseEdge3(False, False)
    if kip:
        edges_use = sw_model.GetActiveSketch2.GetSketchSegments[-1]
        edges_use.ConstructionGeometry = True
        edges_use.Select4(True, selection_data)
        sw_model.SketchManager.SketchOffset2(0.001, False, True, 0, 0, True)

    # create cut extrude
    sw_model.FeatureManager.FeatureCut4(False, False, False, 0, 0, 0.010, 0.010, False, False, False,
                                        False, 0, 0, False, False, False, False, False, True, True, True,
                                        True, False, 0, 0, False, False)
    sw_model.ClearSelection2(True)
    sw_model.EditAssembly()
    a = sw_model.EditRebuild3
    sw_app.SendmsgToUser('Отверстие успешно создано')
    print('Отверстие успешно создано')


def cut_extrude():
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
    create_cut_extrude(sw_app, sw_model, kip=False)


def cut_extrude_kip():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
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
    create_cut_extrude(sw_app, sw_model, kip=True)
