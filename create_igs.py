import os
import sys
import asyncio
import win32com.client


def create_dir(directory: str) -> tuple[str, str]:
    """Create or clear directory"""
    list_directory: list = directory.split('\\')
    list_directory.insert(-1, 'IGS')
    file_name = list_directory.pop()
    new_directory: str = '\\'.join(list_directory)
    if os.path.isdir(new_directory):
        for file in os.listdir(new_directory):
            if file.split('.')[-1] == 'IGS':
                os.remove(new_directory + '\\' + file)
        else:
            print('Директория была очищена от IGS файлов')
    try:
        os.makedirs(new_directory)
    except FileExistsError:
        print('Директория уже существует')
    return new_directory, file_name


async def main():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    sw_model = sw_app.ActiveDoc
    directory, extension = sw_model.GetPathName.split('.')
    if sw_model.GetType != 1:
        print('Активна не деталь')
        return

    # create dir
    new_directory, file_name = create_dir(directory)

    # delete not used configurations
    configs: tuple = sw_model.GetConfigurationNames
    sw_model.ShowConfiguration2(configs[0])
    for name_config in configs:
        sw_model.DeleteConfiguration2(name_config)
    configs: tuple = sw_model.GetConfigurationNames
    sw_model.ShowConfiguration2(configs[1])
    sw_model.DeleteConfiguration2(configs[0])

    # save configurations
    configs: tuple = sw_model.GetConfigurationNames
    for name_config in configs:
        sw_model.ShowConfiguration2(name_config)
        sw_model.DeleteConfiguration2(name_config)
        new_name: str = new_directory + '\\' + file_name + " - " + name_config + ".IGS"
        sw_model.SaveAs3(new_name, 0, 2)


try:
    if __name__ == "__main__":
        asyncio.run(main())
except KeyboardInterrupt:
    sys.exit()
