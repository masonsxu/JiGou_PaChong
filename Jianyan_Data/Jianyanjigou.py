import time #控制代码运行时间
import os #控制文件
import pytesseract #识别引擎
import requests #网页
import json #处理json数据
import pandas as pd #处理excel
from openpyxl import load_workbook #处理excel
from selenium import webdriver #控制浏览器进行自动化登陆
from PIL import Image #处理图片
from collections import defaultdict # 初始化像素字典为
from bs4 import BeautifulSoup #处理网页源代码

#模拟谷歌浏览器
def get_HTML(url):
    header = {
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/68.0.3440.106 Safari/537.36',
    }  # 伪装header
    # 捕捉异常
    try:
        r = requests.get(url, timeout=30, headers=header)
        r.raise_for_status()
        r.encoding = r.apparent_encoding  # 指定编码形式
        return r.text
    except:
        return "please inspect your url or setup"

# 写入excel文件中多个sheet表单中
def excelAddSheet(dataframe, outfile, name):
    writer = pd.ExcelWriter(outfile, enginge='openpyxl')  # 设置文件路径和引擎
    if os.path.exists(outfile) is not True:  # 判断文件内是否为空
        dataframe.to_excel(writer, name, index=None)
    else:
        book = load_workbook(writer.path)
        writer.book = book
        dataframe.to_excel(excel_writer=writer, sheet_name=name, index=None)  # 进行写入操作，设置sheet name,无索引
    writer.save()
    print('save - Successful!!!')
    writer.close()
    print('close - Successful!!!')

# 解析目标网页的html
URL_list_Accpet = []  # 定义list保存url
URL_list_Power = []  # 定义list保存url
datalist = []  # 定义list保存当前页面内的数据
def get_information_from_url(url):
    text = get_HTML(url)  # 调用get_HTML方法获得网页源代码
    soup = BeautifulSoup(text, "html.parser")  # 解析text中的HTML
    table_datas = soup.find_all('table')  # 获取所有的table标签内容
    url_list_power = []  # 定义list保存url
    url_list_accpet = []  # 定义list保存url
    if table_datas[1] != None:
        for table_url in table_datas[1].find_all('a'):
            str1 = str(table_url)
            str2 = 'https://las.cnas.org.cn/LAS/publish/queryPublishSignatory.action?baseinfoId=' + str1[138:170]
            str3 = 'https://las.cnas.org.cn/LAS/publish/queryPublishIBAbilityQuery.action?baseinfoId=' + str1[138:170]
            if len(str2) >= 108:
                url_list_power.append(str2)  # 将url加入到相应的list中
                url_list_accpet.append(str3)  # 将url加入到相应的list中


    data = []
    div_data = [div.get_text().strip().replace('\r', '').replace('\t', '').replace('\n', '').replace('\x1b', '')
                for div in soup.find_all('div', class_='T1')]  # 解析网页源代码并剔除制表符

    div_str = str(div_data)  # 将网页源代码的数据转换为str型
    url_list_power.append(div_str[2:-2])  # 截取字符串
    url_list_accpet.append(div_str[2:-2])  # 截取字符串

    for colomns_data in div_data:
        data.append(colomns_data)  # 遍历div标签内的数据并加入到data中

    table_data = table_datas[0]
    i = 0
    for tr_data in table_data.find_all('tr'):  # 遍历tr标签内的数据并加入到data中
        i = i + 1
        j = 1
        for span_data in tr_data.find_all('span', class_='clabel'):
            j = j + 1
            span_data = span_data.get_text().strip().replace('\x1b', '')
            data.append(span_data)

        if i == 7 and j < 3:
            data.append('')
    datalist.append(data)

    # 连接两个日期字符串
    span_str = str(data[10]+' - '+data[11]).replace('\r', '').replace('\n', '').replace('\t', '')
    url_list_power.append(span_str)
    url_list_accpet.append(span_str)
    if len(url_list_power[0]) >= 108:  # 判断list中的数据是否符合url
        URL_list_Power.append(url_list_power)
        URL_list_Accpet.append(url_list_accpet)

Data_list = []
def get_information_from_url_Power(url):
    text = get_HTML(url[0])  # 调用get_HTML方法获得网页源代码
    soup = BeautifulSoup(text, "html.parser")  # 解析text中的HTML

    html = soup.get_text()   # 网页源代码转换为字符型
    json_data = json.loads(html)  # 将网页源代码加载为json格式，方便后面操作

    infos = json_data['data']
    lastnum = int(json_data['sizePerPage'])
    for info in infos:
        try:
            num = int(info['num'])
            ss = "正在爬取第%d条信息,共%d条" % (num, lastnum)
            print(ss)
            zuzhi_name = str(url[1]).replace('\r', '').replace('\n', '').replace('\t', '')
            date_str = str(url[2]).replace('\r', '').replace('\n', '').replace('\t', '')
            data_list = []
            data_list.append(zuzhi_name)
            data_list.append(date_str)
            # 姓名
            try:
                user_name = info['nameCh'].strip().replace('\r', '')\
                    .replace('\n', '').replace('\t', '').replace('\x1b', '')

            except:
                user_name = ""
            data_list.append(user_name)

            # 授权签字领域
            try:
                comment_date = info['authorizedFieldCh'].strip()\
                    .replace('\r', '').replace('\n', '').replace('\t', '').replace('\x1b', '')

            except:
                comment_date = ""
            data_list.append(comment_date)
            Data_list.append(data_list)
        except:
            sss = '爬取第%d信息失败,跳过爬取' % (num)
            print(sss)
            pass


Data_list_Accept = []
def get_information_from_url_Accept(url):
    text = get_HTML(url[0])  # 调用get_HTML方法获得网页源代码
    soup = BeautifulSoup(text, "html.parser")  # 解析text中的HTML
    html = soup.get_text()  # 将网页源代码转换为字符型
    json_data = json.loads(html)  # 将json网页源代码转换为json型，方便后面操作

    infos = json_data['data']
    for info in infos:
        try:
            zuzhi_name = str(url[1]).strip()  # 注意连接前后的数据
            date_str = str(url[2]).strip()  # 注意连接前后的数据
            data_list = []
            data_list.append(zuzhi_name)
            data_list.append(date_str)
            # 分组名称
            try:
                typeName = info['typeName']

            except:
                typeName = ""
            if typeName == None:
                typeName = '暂无分组'
            typeName.strip()
            data_list.append(typeName)

            # 检验对象序号
            try:
                num = info['num']

            except:
                num = ""
            data_list.append(str(num))

            #检验对象
            try:
                fieldch = info['fieldch'].strip().replace('\r','').replace('\n','').replace('\t','').replace('\x1b','')

            except:
                fieldch = ''
            data_list.append(fieldch)

            #检验项目名称序号
            try:
                detnum = info['detnum']

            except:
                detnum =''
            data_list.append(str(detnum))

            # 检验项目名称
            try:
                descriptCh = info['descriptCh'].strip().replace('\r','').replace('\n','').replace('\t','').replace('\x1b','')

            except:
                descriptCh = ""
            data_list.append(descriptCh)

            # 检验标准序号
            try:
                stdNum = info['stdNum']

            except:
                stdNum = ""
            data_list.append(str(stdNum))

            # 依据的检测标准
            try:
                standardCh = info['standardCh'].strip().replace('\r','').replace('\n','').replace('\t','').replace('\x1b','')

            except:
                standardCh = ""
            data_list.append(standardCh)

            # 依据的检测标准编号
            try:
                order = info['order'].strip()
            except:
                order = ""
            data_list.append(order)

            #说明
            try:
                restrictCh = info['restrictCh'].strip().replace('\r','').replace('\n','').replace('\t','').replace('\x1b','')

            except:
                restrictCh = ''
            data_list.append(restrictCh)

            Data_list_Accept.append(data_list)

        except:
            sss = '本次爬取信息失败'
            print(sss)
            pass

# 获取图片中像素点数量最多的像素
def get_threshold(image):
    pixel_dict = defaultdict(int) #初始化像素字典为0

    # 像素及该像素出现次数的字典
    rows, cols = image.size
    for i in range(rows):
        for j in range(cols):
            pixel = image.getpixel((i, j))
            pixel_dict[pixel] += 1

    count_max = max(pixel_dict.values())  # 获取像素出现出多的次数
    pixel_dict_reverse = {v: k for k, v in pixel_dict.items()}
    threshold = pixel_dict_reverse[count_max]  # 获取出现次数最多的像素点

    return threshold


# 按照阈值进行二值化处理
# threshold: 像素阈值
def get_bin_table(threshold):
    # 获取灰度转二值的映射table
    table = []
    for i in range(256):
        rate = 0.1  # 在threshold的适当范围内进行处理
        if threshold * (1 - rate) <= i <= threshold * (1 + rate):
            table.append(1)
        else:
            table.append(0)
    return table


# 去掉二值化处理后的图片中的噪声点
def cut_noise(image):
    rows, cols = image.size  # 图片的宽度和高度
    change_pos = []  # 记录噪声点位置

    # 遍历图片中的每个点，除掉边缘
    for i in range(1, rows - 1):
        for j in range(1, cols - 1):
            # pixel_set用来记录该店附近的黑色像素的数量
            pixel_set = []
            # 取该点的邻域为以该点为中心的九宫格
            for m in range(i - 1, i + 2):
                for n in range(j - 1, j + 2):
                    if image.getpixel((m, n)) != 1:  # 1为白色,0位黑色
                        pixel_set.append(image.getpixel((m, n)))

            # 如果该位置的九宫内的黑色数量小于等于4，则判断为噪声
            if len(pixel_set) <= 4:
                change_pos.append((i, j))

    # 对相应位置进行像素修改，将噪声处的像素置为1（白色）
    for pos in change_pos:
        image.putpixel(pos, 1)

    return image  # 返回修改后的图片


# 识别图片中的数字加字母
# 传入参数为图片路径，返回结果为：识别结果
def OCR_lmj(img_path):
    image = Image.open(img_path)  # 打开图片文件
    imgry = image.convert('L')  # 转化为灰度图

    # 获取图片中的出现次数最多的像素，即为该图片的背景
    max_pixel = get_threshold(imgry)

    # 将图片进行二值化处理；
    table = get_bin_table(threshold=max_pixel)
    out = imgry.point(table, '1')

    # 去掉图片中的噪声（孤立点）
    out = cut_noise(out)
    out_image = out.crop((0, 0, 165, 60))
    # 保存图片
    out_image.save('E:\\Pycharm_xjf\\JiGou_PaChong\\Image\\text.png')

    # 识别图片中的数字和字母
    text = pytesseract.image_to_string(out_image)

    return text

#使用selenium对浏览器进行自动化测试
driver = webdriver.Chrome()
URL_list = []
def connent_url(location):
    time.sleep(3)  # 让程序暂停3秒防止过快导致的错误
    url = 'https://las.cnas.org.cn/LAS/publish/externalQueryIB.jsp'
    driver.get(url)

    orgAddress = driver.find_elements_by_id('orgAddress')
    orgAddress[0].send_keys(location)

    btn = driver.find_element_by_class_name('btn')
    btn.click()
    time.sleep(6) # 让程序暂停6秒防止过快导致的错误

    accept = False
    while not accept:
        try:

            # # 先对浏览器当前页面截图，并存储
            # image = driver.find_element_by_id('pirlAuthInterceptIframe')
            # # 保存截取后的图片
            # image.screenshot('E:\\Pycharm_xjf\\JiGou_PaChong\\Image\\Datascreenshot.png')
            # print('截图OK!!!')
            # time.sleep(3)  # 防止由于网速，可能图片还没保存好，就开始识别
            #
            # # 调用OCR_lmj方法，变量code_str即为识别出的图片数字str类型
            # code_str = OCR_lmj('E:\\Pycharm_xjf\\JiGou_PaChong\\Image\\Datascreenshot.png')
            # print('code：' + code_str)

            pirlbutton1 = driver.find_element_by_xpath('//*[@id="pirlbutton1"]')
            print('Login')
            pirlbutton1.click()
            time.sleep(5)

            flag_str = driver.find_element_by_id('pirlAuthInterceptDiv_c').is_displayed()
            print(flag_str)

            if not flag_str:
                break

        #     # 判断识别后的验证码长度是否是4，不是的话模拟点击更换验证码
        #     if len(code_str) == 4:
        #         authInterceptVName = driver.find_element_by_xpath('//*[@id="authInterceptVName"]')
        #         authInterceptVName.send_keys(code_str)
        #
        #         # 模拟点击确认
        #         pirlbutton1 = driver.find_element_by_xpath('//*[@id="pirlbutton1"]')
        #         print('Login')
        #         pirlbutton1.click()
        #         time.sleep(3)
        #
        #         # 获取验证码弹窗是否还可见
        #         flag_str = driver.find_element_by_id('pirlAuthInterceptDiv_c').is_displayed()
        #         print(flag_str)
        #
        #         # 如果验证码弹窗不存在的话停止循环
        #         if not flag_str:
        #             break
        #     else:
        #         href = driver.find_element_by_xpath('/html/body/div[6]/div[1]/div[2]/table/tbody/tr[1]/td[2]/a')
        #         href.click()
        #         time.sleep(3)
        #
        except ValueError:  # 抛出value值错误
            accept = False
            time.sleep(3)

    # 查看搜索结果
    flag_str = driver.find_element_by_xpath('/html/body/div[4]/div[3]/table/tbody[1]/tr/td/div').is_displayed()
    if not flag_str:
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')  # 解析网页源代码保存结果url

        for url in soup.find_all('a'):
            str1 = str(url)
            str2 = "https://las.cnas.org.cn" + str1[48:117]
            if len(str2) > 78:
                URL_list.append(str2)

if __name__ == '__main__':
    df = pd.read_excel('E:\\Pycharm_xjf\\JiGou_PaChong\\Data\\机构地址.xlsx',sheet_name='Sheet1')
    data = df['机构地址']

    for key in data:
        connent_url(str(key))

    driver.close()#关闭selenium控制的浏览器
    driver.quit()#退出浏览器


    df_Sheet5 = pd.DataFrame(URL_list, columns=['url'])
    df_Sheet5.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Data\\Jigou_url1.csv')
    # i = 1
    # for url in URL_list:
    #     get_information_from_url(url)
    #     print(i)
    #     i = i+1
    #
    # df_Sheet1 = pd.DataFrame(datalist,
    #                          columns=['机构名称', '注册编号', '报告/证书允许使用认可标识的其他名称', '联系人', '联系电话', '邮政编码', '传真号码', '网站地址',
    #                                   '电子邮箱', '单位地址', '认证开始时间', '认证结束时间', '暂停项目/参数', '机构特点'])
    # print("Sheet1 - Successful!!!")
    # df_Sheet1.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Data\\Jigou_Data1.csv')
    # df_Sheet4 = pd.DataFrame(URL_list_Power, columns=['1', '2', '3'])
    # df_Sheet4.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Data\\Jigou_url2.csv')
    # i = 1
    # for url in URL_list_Power:
    #     get_information_from_url_Power(url)
    #     print(i)
    #     i = i+1
    #
    # df_Sheet2 = pd.DataFrame(Data_list, columns=['机构名称', '周期', '姓名', '授权签字领域'])
    # print('Sheet2 - Successful!!!')
    # df_Sheet2.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Data\\Jigou_Data2.csv')
    # df_Sheet6 = pd.DataFrame(URL_list_Accpet, columns=['1', '2', '3'])
    # df_Sheet6.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Data\\Jigou_url3.csv')
    # i = 1
    # for url in URL_list_Accpet:
    #     get_information_from_url_Accept(url)
    #     print(i)
    #     i = i+1
    #
    # df_Sheet3 = pd.DataFrame(Data_list_Accept,
    #                          columns=['机构名称', '周期', '分组名称', '检验对象序号', '检验对象', '检验项目名称序号', '检验项目名称', '依据的检验标准序号',
    #                                   '依据的检测标准', '方法程序及编号', '说明'])
    # print('Sheet3 - Successful!!!')
    # df_Sheet3.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Data\\Jigou_Data3.csv')
    #
    # outfile = 'E:\\Pycharm_xjf\\JiGou_PaChong\\Data\\Jigou_Data.xlsx'
    # excelAddSheet(df_Sheet1, outfile, '机构信息')
    # excelAddSheet(df_Sheet2, outfile, '能力范围')
    # excelAddSheet(df_Sheet3, outfile, '认可的检验能力范围')