import os
import time
import zipfile


if __name__ == '__main__':
    os.system('python3 -m PyInstaller -F main.py')
    main_file = './dist/main.exe'
    while True:
        if os.path.exists(main_file):
            zip_file = zipfile.ZipFile('dist/淘宝订单转erp.zip', 'w')
            zip_file.write(main_file, '淘宝订单转erp.exe')
            zip_file.close()
            os.unlink(main_file)
            break
        time.sleep(2)
