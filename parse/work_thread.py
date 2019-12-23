import threading
from parse.parse import Parse


class ParseThread(threading.Thread):

    def __init__(self, filename: str, save_name: str, shop_code: str, password: str, cb):
        threading.Thread.__init__(self)
        self.filename = filename
        self.save_name = save_name
        self.shop_code = shop_code
        self.cb = cb
        self.password = password

    def run(self) -> None:
        try:
            parse = Parse(self.filename, self.save_name, self.shop_code, self.password, self.cb)
            parse.do_parse()
            self.callback('处理完毕', 'done')
        except Exception as e:
            self.callback(e, 'error')

    def callback(self, message, state: str):
        if callable(self.cb):
            self.cb(message, state)
