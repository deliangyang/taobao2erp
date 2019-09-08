import wx
from ui.windows import MainWindows

if __name__ == '__main__':
    app = wx.App(False)
    frame = MainWindows(None, "淘宝订单erp")
    app.MainLoop()
