from functions.archive.create_drill_sheet import get_ready
from functions.general_functions import create_app_model, check_edge, check_surface, create_select_man_data, \
    add_unique_conf, check_assembly


def save_edges_surface(sel_manager, count_select):
    """save select objects"""

    edges: list = []
    for i in range(1, count_select):
        edges.append(sel_manager.GetSelectedObject6(i, -1))
    pipe = sel_manager.GetSelectedObjectsComponent4(count_select, -1)
    return edges, pipe


def create_conf(sw_model, pipe, sel_data):
    """create conf in element"""

    name_conf: str = pipe.ReferencedConfiguration
    if not name_conf.endswith(')'):
        pipe.Select4(False, sel_data, False)
        sw_model.EditPart()
        sw_model_pipe = sw_model.GetEditTarget
        all_name_conf: tuple = sw_model_pipe.GetConfigurationNames

        name_new_conf = add_unique_conf(sw_model_pipe, name_conf, all_name_conf)

        sw_model.EditAssembly()
        pipe.ReferencedConfiguration = name_new_conf
        sw_model.ClearSelection2(True)
        return


def create_plane(sw_model, pipe, edge, sel_data):
    """create plane for edge"""

    pipe.Select4(False, sel_data, False)
    sw_model.EditPart()
    sw_model.ClearSelection2(True)
    edge.Select4(False, sel_data)
    if get_ready():
        return
    plane_edge = sw_model.FeatureManager.InsertRefPlane(4, 0, 0, 0, 0, 0)
    sw_model.ClearSelection2(True)
    if not plane_edge:
        return
    return plane_edge


def create_sketch(sw_model, plane_edge, edges, sel_data) -> bool:
    """create sketch and transfer"""

    plane_edge.Select4(False, sel_data)
    sw_model.SketchManager.InsertSketch(True)
    for edge in edges:
        edge.Select4(True, sel_data)
    sketch = sw_model.SketchManager.SketchUseEdge3(False, False)
    if sketch:
        return True
    return False


def edit_radius_edges(sw_model, sel_data, pipe):
    """edit radius edges for template"""

    pipe_name: str = pipe.Name2
    pipe_material: str = 'stainless' if 'н_ж' in pipe_name else 'black'

    without_saddle_black: dict = {'Dn 125': ((0.0423,), (0.041, 0.05, 0.069)),
                                  'Dn 150': ((0.0423, 0.048), (0.05, 0.069)),
                                  'Dn 200': ((0.0423, 0.048, 0.057), (0.069,)),
                                  'Dn 250': ((0.0423, 0.048, 0.057), (0.069,)),
                                  'Dn 300': ((0.0423, 0.048, 0.057, 0.076), ())}
    without_saddle_stainless: dict = {'Dn 125 н_ж': ((0.0424,), (0.0443, 0.0423)),
                                      'Dn 150 н_ж': ((0.0424, 0.0483), ())}

    if pipe_material == 'black':
        black = True
        saddle: tuple = (0.081, 0.1, 0.125, 0.151, 0.209, 0.261, 0.313)
        without_saddle: tuple = (0.0213, 0.0268, 0.0335, 0.015, 0.03, 0.0268, 0.0335, 0.04231, 0.048, 0.057)
        maybe_saddle: tuple = (0.0359, 0.041, 0.05, 0.069)
        for key, value in without_saddle_black.items():
            if key in pipe_name:
                total_saddle = saddle + value[1]
                total_without_saddle = without_saddle + value[0]
                break
        else:
            total_saddle = saddle + maybe_saddle
            total_without_saddle = without_saddle

    elif pipe_material == 'stainless':
        black = False
        saddle: tuple = (0.0563, 0.0543, 0.0721, 0.0849, 0.104, 0.1337, 0.153, 0.2131)
        without_saddle: tuple = (0.0213, 0.0269, 0.0337, 0.015, 0.027, 0.03, 0.04, 0.048, 0.056, 0.0654)
        maybe_saddle: tuple = (0.0384, 0.0443, 0.0364, 0.0423)
        for key, value in without_saddle_stainless.items():
            if key in pipe_name:
                total_saddle = saddle + value[1]
                total_without_saddle = without_saddle + value[0]
                break
        else:
            total_saddle = saddle + maybe_saddle
            total_without_saddle = without_saddle

    else:
        return False

    # kip: tuple = (0.0213, 0.0269, 0.0268, 0.0337, 0.335, 0.027, 0.03, 0.04, 0.048, 0.056, 0.0654, 0.0268,
    #               0.0335, 0.048, 0.057, 0.015)
    # black_steel: tuple = (0.0359, 0.041, 0.05, 0.053, 0.069, 0.081, 0.1, 0.125, 0.151, 0.209, 0.261, 0.313, 0.363)
    # stainless_steel: tuple = (0.0384, 0.0443, 0.0563, 0.0721, 0.0849, 0.104, 0.1337, 0.153, 0.2131, 0.0364, 0.0423)

    for edge in sw_model.GetActiveSketch2.GetSketchSegments:
        edge_diameter = round(edge.GetRadius * 2, 4)
        if edge_diameter in total_without_saddle:
            if not sketch_off_set(sw_model, edge, sel_data, 0.001):
                return False
        elif edge_diameter in total_saddle:
            if black:
                continue
            else:
                if not sketch_off_set(sw_model, edge, sel_data, -0.001):
                    return False
        else:
            return False
    return True


def sketch_off_set(sw_model, edge, sel_data, value):
    edge.ConstructionGeometry = True
    edge.Select4(False, sel_data)
    off_set = sw_model.SketchManager.SketchOffset2(value, False, True, 0, 0, True)
    sw_model.ClearSelection2(True)
    return off_set


def create_cut(sw_model):
    """create cut extrude in two variants"""

    feature_cut = create_feature_cut(sw_model.FeatureManager, False)
    if not feature_cut:
        feature_cut = create_feature_cut(sw_model.FeatureManager, True)
    sw_model.ClearSelection2(True)
    sw_model.EditAssembly()
    a = sw_model.EditRebuild3
    return feature_cut


def create_feature_cut(feature_manager, arg1):
    feature_cut = feature_manager.FeatureCut4(True, False, arg1, 2, 0, 0, 0, False, False, False, False,
                                              0, 0, False, False, False, False, False, True, True, True, True, False,
                                              0, 0, False, False)
    return True if feature_cut else False


def main_any_cut():
    """initialization SW and main"""

    sw_app, sw_model = create_app_model()

    sel_manager, sel_data = create_select_man_data(sw_model)

    count_select = sel_manager.GetSelectedObjectCount2(-1)

    if not check_assembly(sw_app, sw_model):
        return

    for i in range(1, count_select):
        if not check_edge(sw_app, sel_manager, i):
            return

    if not check_surface(sw_app, sel_manager, count_select):
        return

    edges, pipe = save_edges_surface(sel_manager, count_select)
    sw_model.ClearSelection2(True)

    create_conf(sw_model, pipe, sel_data)

    plane_edge = create_plane(sw_model, pipe, edges[0], sel_data)
    if not plane_edge:
        sw_app.SendmsgToUser('⛔⛔ Плоскость не создалась ⛔⛔')
        return

    if not create_sketch(sw_model, plane_edge, edges, sel_data):
        sw_app.SendmsgToUser('⛔⛔ Эскиз не создался ⛔⛔')
        return

    if not edit_radius_edges(sw_model, sel_data, pipe):
        sw_app.SendmsgToUser('⛔⛔ Неверная кромка ⛔⛔')
        return

    if not create_cut(sw_model):
        sw_app.SendmsgToUser('⛔⛔ Не удалось создать отверстие ⛔⛔')
        return

    sw_app.SendmsgToUser('Отверстия успешно созданы')
