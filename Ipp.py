import os
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler, FileSystemEvent
import IPPRequset
import win32api
import subprocess
import json


def read_json_config(file_path):
    with open(file_path, 'r') as f:
        config = json.load(f)
    return config


class PrintOnFileCreateHandler(FileSystemEventHandler):
    def __init__(self, json_file_path):
        config = read_json_config(json_file_path)
        self.printer_url = config["printer_uri"]
        self.watch_folder = config["watch_folder"]
        self.file_extension = config["file_extension"]
        self.gs_path = config["gs_path"]
        self.host = config["server_host"]

    def on_created(self, event):
        if not event.is_directory and event.src_path.endswith(self.file_extension):
            print(f"find file:{event.src_path}")
            if IPPRequset.create_get_printer_ipp(self.printer_url.encode('utf-8'), self.host) == 200:
                print("打印机获取成功")
                file_path = event.src_path
                pdf_file = self.ps_to_pdf(file_path)
                os.remove(file_path)  # 转换完成删除原ps文件
                filename_with_ext = os.path.basename(event.src_path)
                # 分割文件名和扩展名
                filename, extension = os.path.splitext(filename_with_ext)
                #username = input("请输入您的名字: \r\n")
                username = win32api.GetUserName()
                print("登录的用户为:", username)
                if IPPRequset.print_file(self.printer_url.encode('utf-8'), pdf_file, filename.encode('utf-8'), username.encode('utf-8'), self.host) == 200:
                    print(filename + "文件打印成功,删除文件")
                    os.remove(pdf_file)  # 完成后删除pdf文件
                else:
                    print("文件打印失败")
            else:
                print("打印机获取失败")

    def ps_to_pdf(self, ps_file):
        base_path, file_extension = os.path.splitext(ps_file)  # 分离路径和扩展名
        pdf_file = base_path + ".pdf"  # 拼接新的文件路径和扩展名

        # 构建Ghostscript命令行参数
        args = [self.gs_path,
                '-dNOPAUSE', '-dBATCH', '-sDEVICE=pdfwrite',
                '-sOutputFile=' + pdf_file, ps_file]

        # 调用subprocess执行命令
        subprocess.run(args, check=True)
        return pdf_file


event_handler = PrintOnFileCreateHandler(json_file_path='config.json')
observer = Observer()
observer.schedule(event_handler, event_handler.watch_folder, recursive=False)
observer.start()
try:
    while True:
        time.sleep(1)
except KeyboardInterrupt:
    observer.stop()

observer.join()
