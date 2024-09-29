from functions.general_functions import create_app_model, check_assembly, check_unselect_element, create_select_man_data


def filter_only_pipe(sw_model) -> list:
    """filter only pipes"""

    pipes: list = []
    for i in range(1, sw_model.SelectionManager.GetSelectedObjectCount2(-1) + 1):
        component = sw_model.SelectionManager.GetSelectedObjectsComponent4(i, -1)
        if component.Name2.startswith('Труба'):
            pipes.append(component)
    return pipes


def create_pipe_conf(sw_model, pipes):
    """create in pipe conf"""

    sel_manager, sel_data = create_select_man_data(sw_model)

    for pipe in pipes:
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
            break
    return


def add_unique_conf(sw_model_pipe, name_conf, all_name_conf) -> str:
    """add in pipe unique configurations"""

    for i in range(2, 51):
        name_new_conf: str = f'{name_conf}({i})'
        if name_new_conf not in all_name_conf:
            sw_model_pipe.ConfigurationManager.AddConfiguration2(name_new_conf, '', '', 128,
                                                                 name_conf, '', True)
            return name_new_conf

def main_create_select_conf():
    """initialization SW and main"""

    sw_app, sw_model = create_app_model()

    if not check_assembly(sw_app, sw_model):
        return

    if not check_unselect_element(sw_app, sw_model):
        return

    pipes = filter_only_pipe(sw_model)

    sw_model.ClearSelection2(True)

    if not create_pipe_conf(sw_model, pipes):
        sw_app.SendmsgToUser('Конфигурации успешно добавлены')


main_create_select_conf()
