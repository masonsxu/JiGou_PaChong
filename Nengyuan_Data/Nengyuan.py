# import requests
# import re
# import urllib.request
# import pandas as pd
# from bs4 import BeautifulSoup
# from fake_useragent import UserAgent
#
# # 模拟谷歌浏览器
# def get_html(url):
#     ua = UserAgent(verify_ssl=False)  # 调用UserAgent库生成ua对象
#     header = {
#         'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/web,image/png,*/*;q=0.8',
#         'user-agent': ua.random,
#     }
#     try:
#         r = requests.get(url, timeout=30, headers=header)
#         r.raise_for_status()
#         r.encoding = r.apparent_encoding  # 指定编码形式
#         return r.text
#     except:
#         return "please inspect your url or setup"
#
# URL_list = []
# def get_information_from_url(url):
#     text = get_html(url)
#     soup = BeautifulSoup(text, "html5lib")  # 解析text中的HTML
#
#     ul_data = soup.find_all(href=re.compile("cxzq/"))
#     for li_str in ul_data:
#         if (li_str.get('href'))[-6:] == '.shtml':
#             url_str = 'https://www.cnas.org.cn' + li_str.get('href')
#             URL_list.append(url_str)
#
# pdf_url_list = []
# def get_pdf_file(url):
#     text = get_html(url)
#     soup = BeautifulSoup(text, 'html5lib')  # 解析html
#
#     pdf_url_data = soup.find_all(href=re.compile('/images/'))
#     for pdf_url_str in pdf_url_data:
#         if (pdf_url_str.get('href'))[-4:] == '.pdf' or (pdf_url_str.get('href'))[-4:] == '.doc':
#             url_str = 'https://www.cnas.org.cn' + pdf_url_str.get('href')
#             pdf_url_list.append(url_str)
#
# # 批量下载pdf文件
# def get_pdf(url):
#     file_name = url.split('/')[-1]
#     r = requests.get(url)
#     with open(file_name, 'wb') as f:
#         f.write(r.content)
#
# if __name__ == '__main__':
#     path = 'E:/Pycharm_xjf/JiGou_PaChong/Nengyuan_Data'
#     get_information_from_url('https://www.cnas.org.cn/cxzq/nyzxsys/index.shtml')
#     print(URL_list)
#     df_Sheet1 = pd.DataFrame(URL_list)
#     path1 = path + "/Nengyuan_url1.csv"
#     df_Sheet1.to_csv(path1)
#
#     i = 1
#     for url in URL_list:
#         get_pdf_file(url)
#         print(i)
#         i = i + 1
#     print(pdf_url_list)
#     df_Sheet2 = pd.DataFrame(pdf_url_list)
#     path1 = path + "/Nengyuan_url2.csv"
#     df_Sheet2.to_csv(path1)
#
#     i = 1
#     for url in pdf_url_list:
#         get_pdf(url)
#         print(i)
#         i = i + 1

import os, stat
from win32com.client import Dispatch, constants, gencache

class word2Pdf:
    def __init__(self, word_file_path):
        self.word_file_path = word_file_path
        self.pdf_file_path = os.path.splitext(self.word_file_path)[0] + '.pdf'
        print(self.pdf_file_path)
        self.word = Dispatch('Word.Application')

    def conver(self):
        try:
            print(str(self.word_file_path))
            doc = self.word.Documents.Open(self.word_file_path, ReadOnly=1)
            print(doc)
            doc.ExportAsFixedFormat(self.pdf_file_path,
                                    constants.wdExportFormatPDF,
                                    Item=constants.wdExportDocumentWithMarkup,
                                    CreateBookmarks=constants.wdExportCreateHeadingBookmarks)
        except Exception as e:
            print('转换异常，异常是：{}'.format(e))
        finally:
            self.word.Quit(constants.wdDoNotSaveChanges)

# 当前文件夹中所有的word文件
if __name__ == '__main__':
    word_file_paths = [fn for fn in os.listdir('E:/Pycharm_xjf/JiGou_PaChong/Nengyuan_Data/data')
                       if fn.endswith(('.doc', '.docx'))]
    for word_file_path in word_file_paths:
        word_file_path = 'E:/Pycharm_xjf/JiGou_PaChong/Nengyuan_Data/data/' + word_file_path
        print(word_file_path)
        # 设置文件可以被其他用户写入
        os.chmod(word_file_path, stat.S_IWRITE)
        conver = word2Pdf(word_file_path)
        conver.conver()