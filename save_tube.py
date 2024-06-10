import os
import sys
import asyncio
import win32com.client


def create_dir(path: str):
    """Check and crete path"""
    try:
        os.makedirs(path)
    except FileExistsError:
        print(f'Директория {path} уже существует')


async def save_tube():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 2:
        print('Активна не сборка')
        return

    # Save tube in a separate file
    components = sw_model.GetComponents(True)
    tubes: list = []
    for component in components:
        component_name: str = component.name2.split('-')[0]
        if component_name.startswith('Труба'):
            if component_name not in tubes:
                tubes.append(component_name)
                part = component.GetModelDoc2
                path_list: list = sw_model.GetPathName.split('\\')
                assembly_name: str = path_list.pop().split('.')[0]
                path_list.append('Трубы')
                path_list.append(assembly_name)
                path: str = '\\'.join(path_list)
                create_dir(path)
                part.SaveAs3(path + '\\' + component_name + '.SLDPRT', 0, 8)
    else:
        print('Трубы успешно сохранены')


try:
    if __name__ == "__main__":
        asyncio.run(save_tube())
except KeyboardInterrupt:
    sys.exit()

    # # Get components assembly
    # components = sw_model.GetComponents(True)
    # for component in components:
    #     print(component.ReferencedConfiguration)
    #     part = component.GetModelDoc2
    #     b = part
    #     print(b)
