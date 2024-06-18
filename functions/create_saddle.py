import asyncio

import pythoncom
import win32com.client


async def create_saddle():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    vt_dispatch = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 1:
        print('Активна не деталь')
        return

    # save choose element
    selection_manager = sw_model.SelectionManager
    if selection_manager.GetSelectedObjectType3(1, -1) != 1:
        print('Не выбрана кромка под седло')
        return
    edges = selection_manager.GetSelectedObject6(1, -1)
    edges_coordinate = selection_manager.GetSelectionPoint2(1, -1)
    sw_model.ClearSelection2(True)

    # create circle
    sw_model.Extension.SelectByID2('Спереди', 'PLANE', 0, 0, 0, False, 0, vt_dispatch, 0)
    sw_model.SketchManager.InsertSketch(True)
    center_circle = -0.2 if edges_coordinate[0] < 0.001 else 2.2
    circle = sw_model.SketchManager.CreateCircle(center_circle, 0, 0, 0.02, 0.02, 0)
    diameter = sw_model.AddDiameterDimension2(0, 0, 0)
    diameter_size = diameter.GetDimension2(0)
    diameter_size.SetSystemValue3(diameter_size.GetSystemValue3(1)[0] + 0.005, 1)

    # create horizontal
    sw_model.Extension.SelectByID2('Point1', 'SKETCHPOINT', center_circle, 0, 0, True, 0, vt_dispatch, 0)
    sw_model.Extension.SelectByID2('Point1@Исходная точка', 'EXTSKETCHPOINT', 0, 0, 0, True, 0, vt_dispatch, 0)
    sw_model.SketchAddConstraints('sgHORIZONTALPOINTS2D')

    # Sketch use edge
    selection_data = selection_manager.CreateSelectData
    selection_data.Mark = 1
    edges.Select4(True, selection_data)
    sw_model.SketchManager.SketchUseEdge3(False, False)
    edges_line = sw_model.GetActiveSketch2.GetSketchSegments[-1]
    edges_line.ConstructionGeometry = True
    point_edges = sw_model.GetActiveSketch2.GetSketchPoints2[-1]
    point_edges.Select4(True, selection_data)
    circle.Select4(True, selection_data)
    sw_model.SketchAddConstraints('sgCOINCIDENT')
    sw_model.ClearSelection2(True)

    # Create cut extrude

    sw_model.FeatureManager.FeatureCut4(True, False, False, 6, 0, 0.5, 0.001, False, False, False, False, 0, 0, False,
                                        False, False, False, False, True, True, True, True, False, 0, 0, False, False)
    sw_model.ClearSelection2(True)
    print('Седло успешно создано')

asyncio.run(create_saddle())