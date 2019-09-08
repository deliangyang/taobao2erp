import codecs
import xlwt
import re


class Parse(object):

    def __init__(self, filename: str, save_name: str, shop_code: str):
        self.filename = filename
        self.save_name = save_name
        self.shop_code = shop_code
        self.re_extract_code = re.compile(r'\d{13}')

    def do_parse(self):
        we_need = ['订单编号', '买家实际支付金额', '收货人姓名', '收货地址', '联系手机', '宝贝总数量', '店铺名称', '宝贝标题', ]
        name = ['order', 'real_amount', 'username', 'address', 'mobile', 'quantity', 'shop_name', 'title',
                'province', 'city', 'county']

        erp_names = ['商铺(必填）', '网店订单号(必填）', '订单编号', 'ＥＲＰ码(必填）', '书号(必填）', '书名',
                     '数量（必填）', '订单价格（必填）', '实付金额', '应付金额', '发货折扣',
                     '邮费', '买家昵称', '付款时间', '收货人所在省（必填）', '收货人所在市（必填）',
                     '收货人所在区（必填）', '地址（必填）', '收货人（必填）', '收货人联系方式（必填）',
                     '店铺', ]

        result = []
        with codecs.open(self.filename, 'r', encoding='gb2312') as f:
            count = 0
            meta = {}
            for line in f.readlines():
                if count == 0:
                    index = 0
                    for keyword in line.split(','):
                        kw = self.replace(keyword)
                        if kw in we_need:
                            meta[name[we_need.index(kw)]] = index
                        index += 1
                    count += 1
                else:
                    item = line.split(',')
                    items = list(map(lambda x: self.replace(x), item))
                    temp = {}
                    for n in name:
                        if n in meta:
                            temp[n] = items[meta[n]]
                    if 'address' in temp:
                        address = temp['address'].split(' ')
                        temp['province'] = address[0]
                        temp['city'] = address[1]
                        temp['county'] = address[2]
                    result.append(temp)
        workbook = xlwt.Workbook()
        worksheet = workbook.add_sheet('sheet')
        index = 0
        for erp in erp_names:
            worksheet.write(0, index, erp)
            index += 1

        count = 1
        for erp in result:
            # 懒得一一映射关系了
            worksheet.write(count, 0, erp['shop_name'])
            worksheet.write(count, 1, erp['order'])
            worksheet.write(count, 2, erp['order'])
            title = self.title_parse(erp['title'])
            erp_code, ok = self.extract_erp_code(title)
            if ok:
                worksheet.write(count, 3, '%s0101' % erp_code)
                worksheet.write(count, 4, erp_code)
            worksheet.write(count, 5, title)
            worksheet.write(count, 6, erp['quantity'])
            worksheet.write(count, 7, '%.2f' % (float(erp['real_amount']) / float(erp['quantity'])))
            worksheet.write(count, 8, erp['real_amount'])
            worksheet.write(count, 9, erp['real_amount'])
            worksheet.write(count, 14, erp['province'])
            worksheet.write(count, 15, erp['city'])
            worksheet.write(count, 16, erp['county'])
            worksheet.write(count, 17, erp['address'])
            worksheet.write(count, 18, erp['username'])
            worksheet.write(count, 19, erp['mobile'].strip("'"))
            worksheet.write(count, 20, self.shop_code)
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
