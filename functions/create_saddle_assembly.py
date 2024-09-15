import pythoncom
import win32com.client

from functions.create_drill_sheet import get_ready


def create_saddle_assembly(sw_app, sw_model, vt_dispatch, plane):
    """Create sketch and FeatureCut"""
    selection_manager_assembly = sw_model.SelectionManager
    selection_data_assembly = selection_manager_assembly.CreateSelectData
    edges_assembly = selection_manager_assembly.GetSelectedObject6(1, -1)
    configuration_assembly = selection_manager_assembly.GetSelectedObjectsComponent4(1, -1)
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
        configuration_assembly.ReferencedConfiguration = name_new_configurations
        configuration_assembly.Select4(True, selection_data_assembly, False)
        sw_model.EditPart()
        if get_ready():
            return

    # create sketch
    part_name: str = configuration_assembly.Name2
    assembly_name: str = sw_model.GetPathName.split('\\')[-1].split('.')[0]
    sw_model.Extension.SelectByID2(f'{plane}@{part_name}@{assembly_name}', 'PLANE', 0, 0, 0,
                                   False, 0, vt_dispatch, 0)
    sw_model.SketchManager.InsertSketch(True)
    sw_model.ClearSelection2(True)

    # sketch use edge
    edges_assembly.Select4(True, selection_data_assembly)
    sw_model.SketchManager.SketchUseEdge3(False, False)
    edges_line = sw_model.GetActiveSketch2.GetSketchSegments[-1]
    edges_line.ConstructionGeometry = True
    point_edges = sw_model.GetActiveSketch2.GetSketchPoints2[-1]
    edges_coordinate: float = point_edges.X

    # create circle
    center_circle = -0.2 if edges_coordinate < 0.001 else 2.2
    sw_model.SketchManager.CreateCircle(center_circle, 0, 0, 0.02, 0.02, 0)
    diameter = sw_model.AddDiameterDimension2(0, 0, 0)
    diameter_size = diameter.GetDimension2(0)
    diameter_size.SetSystemValue3(diameter_size.GetSystemValue3(1)[0] + 0.005, 1)

    # create horizontal
    sw_model.Extension.SelectByID2(f'Point1@{part_name}@{assembly_name}', 'SKETCHPOINT', center_circle, 0, 0, True, 0,
                                   vt_dispatch, 0)
    sw_model.Extension.SelectByID2(f'Point1@Исходная точка@{part_name}@{assembly_name}', 'EXTSKETCHPOINT', 0, 0, 0,
                                   True, 0, vt_dispatch, 0)
    sw_model.SketchAddConstraints('sgHORIZONTALPOINTS2D')
    sw_model.ClearSelection2(True)

    # Coincident circle and edges
    selection_manager = sw_model.SelectionManager
    selection_data = selection_manager.CreateSelectData
    selection_data.Mark = 2
    point_edges.Select4(True, selection_data)
    sw_model.GetActiveSketch2.GetSketchSegments[-2].Select4(True, selection_data)
    sw_model.SketchAddConstraints('sgCOINCIDENT')
    sw_model.ClearSelection2(True)

    # create cut extrude
    feature_cut = sw_model.FeatureManager.FeatureCut4(True, False, False, 6, 0, 0.5, 0.001, False, False, False, False,
                                                      0, 0, False, False, False, False, False, True, True, True, True,
                                                      False, 0, 0, False, False)
    sw_model.ClearSelection2(True)

    # suppression other configurations
    sketch_feature_cut = feature_cut.GetParents[0]
    a = sketch_feature_cut.SetSuppression2(0, 2)
    sw_model.Extension.SelectByID2(f'{feature_cut.Name}@{part_name}@{assembly_name}', 'BODYFEATURE', 0, 0, 0, False, 0,
                                   vt_dispatch, 0)
    a = sw_model.EditUnsuppress2
    sw_model.EditAssembly()
    a = sw_model.EditRebuild3
    sw_app.SendmsgToUser('Седло успешно создано')
    print('Седло успешно создано')


def assembly_saddle_front():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    vt_dispatch = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 2:
        sw_app.SendmsgToUser('Активна не сборка')
        print('Активна не сборка')
        return
    if sw_model.SelectionManager.GetSelectedObjectType3(1, -1) != 1:
        sw_app.SendmsgToUser('Не выбрана кромка под седло')
        print('Не выбрана кромка под седло')
        return
    if sw_model.SelectionManager.GetSelectedObjectCount2(-1) > 1:
        sw_app.SendmsgToUser('Выбрано два объекта')
        print('Выбрано два объекта')
        return
    create_saddle_assembly(sw_app, sw_model, vt_dispatch, plane='Спереди')


def assembly_saddle_above():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    vt_dispatch = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 2:
        sw_app.SendmsgToUser('Активна не сборка')
        print('Активна не сборка')
        return
    if sw_model.SelectionManager.GetSelectedObjectType3(1, -1) != 1:
        sw_app.SendmsgToUser('Не выбрана кромка под седло')
        print('Не выбрана кромка под седло')
        return
    if sw_model.SelectionManager.GetSelectedObjectCount2(-1) > 1:
        sw_app.SendmsgToUser('Выбрано два объекта')
        print('Выбрано два объекта')
        return
    create_saddle_assembly(sw_app, sw_model, vt_dispatch, plane='Сверху')
