import codecs

import xlwt
import os
import re
import openpyxl
import win32com.client


class Parse(object):

    def __init__(self, filename: str, save_name: str, shop_code: str, password: str, callback):
        self.filename = filename
        self.save_name = save_name
        self.shop_code = shop_code
        self.password = password
        self.callback = callback
        self.re_extract_code = re.compile(r'\d{13}')
        self.we_need = ['订单编号', '买家实际支付金额', '收货人姓名', '收货地址', '联系手机', '宝贝总数量', '店铺名称', '宝贝标题', ]
        self.name = ['order', 'real_amount', 'username', 'address', 'mobile', 'quantity', 'shop_name', 'title',
                'province', 'city', 'county']
        self.warning = xlwt.easyxf('pattern: pattern solid, fore_colour pink;')

        self.key_map = [
            'shop_name', 'order', 'order', 'erp_code0101', 'erp_code', 'title',
            'quantity', 'price', 'real_amount', 'real_amount',
            '', '', '', '', 'province', 'city', 'county', 'address',
            'username', 'mobile', 'shop_code',
        ]

    def read_excel(self, filename):
        try:
            app = win32com.client.Dispatch("Excel.Application")
            workbook = app.Workbooks.Open(filename, False, True, None, Password=self.password)
            tmp_filename = r"%s%stmp.csv" % (os.getcwd(), os.sep)
            app.ActiveWorkbook.SaveAs(tmp_filename, 62, "", "")
        except Exception as e:
            raise e
        result = self.read_csv(tmp_filename, 'utf8')
        try:
            os.unlink(tmp_filename)
        except Exception as e:
            print(e)
        return result

    def read_csv(self, filename, encoding):
        result = []
        with codecs.open(filename, 'r', encoding=encoding) as f:
            count = 0
            meta = {}
            for line in f.readlines():
                if count == 0:
                    index = 0
                    for keyword in line.split(','):
                        kw = self.replace(keyword)
                        if kw in self.we_need:
                            meta[self.name[self.we_need.index(kw)]] = index
                        index += 1
                    count += 1
                else:
                    item = line.split(',')
                    items = list(map(lambda x: self.replace(x), item))
                    temp = {}
                    for n in self.name:
                        if n in meta:
                            temp[n] = items[meta[n]]
                    if 'address' in temp:
                        address = temp['address'].split(' ')
                        temp['province'] = address[0]
                        temp['city'] = address[1]
                        temp['county'] = address[2]
                    result.append(temp)
        return result

    def read(self, filename, encoding):
        if filename.endswith('.csv'):
            return self.read_csv(filename, encoding)
        else:
            return self.read_excel(filename)

    def do_parse(self):
        erp_names = ['商铺(必填）', '网店订单号(必填）', '订单编号', 'ＥＲＰ码(必填）', '书号(必填）', '书名',
                     '数量（必填）', '订单价格（必填）', '实付金额', '应付金额', '发货折扣',
                     '邮费', '买家昵称', '付款时间', '收货人所在省（必填）', '收货人所在市（必填）',
                     '收货人所在区（必填）', '地址（必填）', '收货人（必填）', '收货人联系方式（必填）',
                     '店铺', ]

        result = []
        try:
            result = self.read(self.filename, 'gb2312')
        except Exception as e:
            try:
                result = self.read(self.filename, 'gb18030')
            except Exception as e:
                result = self.read(self.filename, 'utf8')

        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet('sheet')
        index = 0
        for erp in erp_names:
            worksheet.write(0, index, erp)
            index += 1

        count = 1
        for erp in result:
            title = self.title_parse(erp['title'])
            titles = title.split('，')
            style = xlwt.Style.default_style
            if len(titles) > 1:
                style = self.warning
            for (index, title) in enumerate(titles):
                # 懒得一一映射关系了
                erp['title'] = title
                erp['shop_code'] = self.shop_code
                erp['mobile'] = erp['mobile'].strip("'")
                erp['price'] = '%.2f' % (float(erp['real_amount']) / float(erp['quantity']))
                erp_code, ok = self.extract_erp_code(title)
                if ok:
                    erp['erp_code0101'] = '%s0101' % erp_code
                    erp['erp_code'] = erp_code
                for (idx, key) in enumerate(self.key_map):
                    content = ''
                    if key and key in erp:
                        content = erp[key]
                    worksheet.write(count, idx, content, style)
                count += 1
        workbook.save(self.save_name)

    @classmethod
    def title_parse(cls, title: str):
        return title.strip('正版现货包邮')\
            .strip('正版现货')\
            .strip('cq')\
            .strip('出版社直发')\
            .strip(' ')

    @classmethod
    def replace(cls, word: str):
        return word.strip('=').strip('"').strip(' ')

    def extract_erp_code(self, title: str) -> (str, bool):
        match = self.re_extract_code.findall(title)
        if match:
            return match[0], True
        return None, False
