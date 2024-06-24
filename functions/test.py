import pythoncom
import win32com.client


def part_saddle_1():
    sw_app = win32com.client.dynamic.Dispatch('SldWorks.Application')
    vt_dispatch = win32com.client.VARIANT(pythoncom.VT_DISPATCH, None)
    sw_model = sw_app.ActiveDoc
    if sw_model.GetType != 1:
        sw_app.SendmsgToUser('Активна не деталь')
        print('Активна не деталь')
        return
    a = sw_model.FeatureByName('Бобышка-Вытянуть2')
    print(a)
    a.SetSuppression2(0, 1, ('100 мм', ))


part_saddle_1()
