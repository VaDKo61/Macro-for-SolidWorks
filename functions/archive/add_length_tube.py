import win32com.client

from functions.archive.create_drill_sheet import get_ready


def add_length_tube(sw_model):
    """add size in select view"""
    view = sw_model.ActiveDrawingView

    # selection manager
    selection_manager = sw_model.SelectionManager
    selection_data = selection_manager.CreateSelectData

    # select tube
    components = view.GetVisibleDrawingComponents
    name_tube: tuple = ('Труба', 'Ниппель', 'Резьба')
    if get_ready():
        return
    for component in components:
        name: str = component.Name.split('/')[-1]
        if name.split()[0] in name_tube:
            component.Select(True, selection_data)

    # add size
    sw_model.InsertModelAnnotations3(2, 32776, False, True, False, True)


def length_tube():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 3:
        sw_app.SendmsgToUser('Активен не чертеж')
        return
    add_length_tube(sw_model)
    sw_app.SendmsgToUser('Размеры труб проставлены')
