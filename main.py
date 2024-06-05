import sys
import asyncio
import win32com.client


async def main():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    sw_model = sw_app.ActiveDoc
    file_name: str = sw_model.GetPathName.split('.')[0]
    configs: tuple = sw_model.GetConfigurationNames
    for i in configs:
        sw_model.ShowConfiguration2(i)
        new_name: str = file_name + " - " + i + ".IGS"
        sw_model.SaveAs3(new_name, 0, 2)


try:
    if __name__ == "__main__":
        asyncio.run(main())
except KeyboardInterrupt:
    sys.exit()
