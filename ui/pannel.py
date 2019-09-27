import wx
import os
import sys
import datetime
from parse.work_thread import ParseThread


class MainPanel(wx.Panel):

    def __init__(self, parent):
        wx.Panel.__init__(self, parent, -1)
        self.status_bar = parent.status_bar
        self.filename = ''
        self.dirname = ''

        text = """
请从淘宝导出指定日期的订单，然后下载导出的订单（CSV）
然后点击按钮"生成ERP订单"，选择淘宝订单，生成ERP订单
        """
        self.st_tips = wx.StaticText(self, 0, label=text, style=wx.TE_LEFT)
        top_box_sizer = wx.BoxSizer(wx.VERTICAL)
        top_box_sizer.Add(self.st_tips, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)

        b_sizer_all = wx.BoxSizer(wx.VERTICAL)

        center_box_sizer = wx.BoxSizer(wx.HORIZONTAL)
        btn_box_sizer = wx.BoxSizer(wx.HORIZONTAL)
        shop_label = wx.StaticText(self, 0, "店铺:", style=wx.TE_LEFT)
        center_box_sizer.Add(shop_label, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)

        self.btn = wx.Button(self, 1, "生成ERP订单")
        self.shop_code = wx.TextCtrl(self, 1, style=wx.TE_LEFT, value='38B71E5310DF46F08360D8BAC4E32E54')

        center_box_sizer.Add(self.shop_code, proportion=1, flag=wx.ALL | wx.ALIGN_CENTER_VERTICAL,
                             border=5)
        btn_box_sizer.Add(self.btn, proportion=0, flag=wx.ALL | wx.CENTER, border=5)

        b_sizer_all.Add(top_box_sizer, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        b_sizer_all.Add(center_box_sizer, proportion=0, flag=wx.EXPAND | wx.ALL, border=5)
        b_sizer_all.Add(btn_box_sizer, proportion=0, flag=wx.CENTER, border=20)
        self.SetSizer(b_sizer_all)
        self.Bind(wx.EVT_BUTTON, self.on_get_file, self.btn)

    def on_get_file(self, e):
        dlg = wx.FileDialog(self, "Choose a file", self.dirname, "", "*.csv", wx.FD_OPEN)
        if dlg.ShowModal() == wx.ID_CANCEL:
            dlg.Destroy()
            return
        self.filename = dlg.GetFilename()
        self.dirname = dlg.GetDirectory()
        dlg.Destroy()
        self.btn.Disable()
        self.init_status_bar()
        self.status_bar.SetStatusText(u"状态：处理中...", 0)
        thread = ParseThread(
            self.dirname + os.sep + self.filename,
            self.get_storage_path(),
            self.shop_code.GetValue(),
            self.after_parse
        )
        thread.start()

    def get_storage_path(self):
        dirname = os.path.join(
            self.get_desktop_path(),
            'erp'
        )
        if not os.path.exists(dirname):
            os.makedirs(dirname)
        return os.path.join(
            dirname,
            (str(datetime.datetime.now()) + '.xls').replace(' ', '-').replace(':', '')
        )

    def after_parse(self, message: str, state: str):
        self.btn.Enable()
        if state == 'done':
            self.status_bar.SetStatusText(u"状态：处理完毕", 0)
            self.status_bar.SetStatusText(u"文件路径：%s" % self.get_storage_path(), 1)
            dirname = os.path.dirname(self.get_storage_path())
            if sys.platform == 'darwin':
                os.system('open %s' % dirname)
            else:
                os.system('start explorer %s' % dirname)
        else:
            self.status_bar.SetStatusText(u"状态：处理失败 %s" % message, 0)
            self.status_bar.SetStatusText(u"文件路径：", 1)

    @classmethod
    def get_desktop_path(cls):
        return os.path.join(os.path.expanduser("~"), 'Desktop')

    def init_status_bar(self):
        self.status_bar.SetStatusText(u"状态：添加淘宝导出订单", 0)
        self.status_bar.SetStatusText(u"文件路径：", 1)
