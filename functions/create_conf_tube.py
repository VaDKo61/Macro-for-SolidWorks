import win32com.client


def create_conf_tube(sw_model, component_name: str, name_conf: str, count: int):
    """create configurations tube"""

    sw_model.EditPart()
    sw_model_tube = sw_model.GetEditTarget
    if name_conf != sw_model_tube.ConfigurationManager.ActiveConfiguration.Name:
        return False
    name_new_conf: str = f'{name_conf}({count})'
    sw_model_tube.ConfigurationManager.AddConfiguration2(name_new_conf, '', '', 128, name_conf, '', True)
    sw_model.EditAssembly()
    return name_new_conf


def search_tube(sw_model):
    """iterate_components assembly and search tube"""

    selection_manager = sw_model.SelectionManager
    selection_data = selection_manager.CreateSelectData
    tubes: dict[str, dict[str, int]] = {}
    for component in sw_model.GetComponents(True):
        if not component.name2.startswith('Труба'):
            continue
        component_name: str = component.name2.split('-')[0]
        conf = component.ReferencedConfiguration
        if not tubes.get(component_name):
            tubes[component_name] = {conf: 1}
        else:
            tubes[component_name][conf] = tubes[component_name].setdefault(conf, 0) + 1
        component.Select4(False, selection_data, False)
        name_new_conf = create_conf_tube(sw_model, component_name, conf, tubes[component_name][conf])
        if not name_new_conf:
            return False
        component.ReferencedConfiguration = name_new_conf
    return True


def initialization_conf_tube():
    """Creates the configuration of tubes if they have the same length"""

    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 2:
        sw_app.SendmsgToUser('Активна не сборка')
        return
    sw_model.ClearSelection2(True)
    if not search_tube(sw_model):
        sw_app.SendmsgToUser('Не удалось добавить конфигурацию')
    else:
        sw_app.SendmsgToUser('Конфигурации не добавлены')
