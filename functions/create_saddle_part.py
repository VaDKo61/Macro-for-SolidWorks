import pythoncom
import win32com.client

from functions.create_drill_sheet import get_ready


def create_saddle_part(sw_app, sw_model, vt_dispatch):
    """Create sketch and FeatureCut"""
    # save choose element
    selection_manager = sw_model.SelectionManager
    if selection_manager.GetSelectedObjectType3(1, -1) != 1:
        sw_app.SendmsgToUser('Не выбрана кромка под седло')
        return
    if selection_manager.GetSelectedObjectCount2(-1) > 1:
        sw_app.SendmsgToUser('Выбрано два объекта')
        return
    edges = selection_manager.GetSelectedObject6(1, -1)
    sw_model.ClearSelection2(True)

    # create configuration
    name_active_configuration: str = sw_model.ConfigurationManager.ActiveConfiguration.Name
    if not name_active_configuration.endswith(')'):
        name_configurations: str = sw_model.GetConfigurationNames
        for i in range(2, 51):
            name_new_configurations: str = f'{name_active_configuration}({i})'
            if name_new_configurations not in name_configurations:
                sw_model.ConfigurationManager.AddConfiguration2(name_new_configurations, '', '', 128,
                                                                name_active_configuration, '', True)
                break
        else:
            sw_app.SendmsgToUser('Не удалось добавить конфигурацию')
            return
        sw_model.ShowConfiguration2(name_new_configurations)

    # create sketch
    if get_ready():
        return
    sw_model.Extension.SelectByID2('Справа', 'PLANE', 0, 0, 0, False, 0, vt_dispatch, 0)
    sw_model.SketchManager.InsertSketch(True)
    sw_model.ClearSelection2(True)

    # sketch use edge
    selection_data = selection_manager.CreateSelectData
    selection_data.Mark = 1
    edges.Select4(True, selection_data)
    sw_model.SketchManager.SketchUseEdge3(False, False)
    edges_line = sw_model.GetActiveSketch2.GetSketchSegments[-1]
    edges_line.ConstructionGeometry = True
    point_edges = sw_model.GetActiveSketch2.GetSketchPoints2[-1]
    edges_coordinate: float = point_edges.X

    # create circle
    center_circle = -0.2 if edges_coordinate < 0.001 else 2.2
    circle = sw_model.SketchManager.CreateCircle(center_circle, 0, 0, 0.02, 0.02, 0)
    diameter = sw_model.AddDiameterDimension2(0, 0, 0)
    diameter_size = diameter.GetDimension2(0)
    diameter_size.SetSystemValue3(diameter_size.GetSystemValue3(1)[0] + 0.005, 1)

    # create horizontal
    sw_model.Extension.SelectByID2('Point1', 'SKETCHPOINT', center_circle, 0, 0, True, 0, vt_dispatch, 0)
    sw_model.Extension.SelectByID2('Point1@Исходная точка', 'EXTSKETCHPOINT', 0, 0, 0, True, 0, vt_dispatch, 0)
    sw_model.SketchAddConstraints('sgHORIZONTALPOINTS2D')
    sw_model.ClearSelection2(True)

    # Coincident circle and edges
    point_edges.Select4(True, selection_data)
    circle.Select4(True, selection_data)
    sw_model.SketchAddConstraints('sgCOINCIDENT')
    sw_model.ClearSelection2(True)

    # create cut extrude
    sw_model.FeatureManager.FeatureCut4(True, False, False, 6, 0, 0.5, 0.001, False, False, False, False, 0, 0, False,
                                        False, False, False, False, True, True, True, True, False, 0, 0, False, False)
    sw_model.ClearSelection2(True)
    sw_app.SendmsgToUser('Седло успешно создано')


def part_saddle():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    vt_dispatch = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 1:
        sw_app.SendmsgToUser('Активна не деталь')
        return
    create_saddle_part(sw_app, sw_model, vt_dispatch)
