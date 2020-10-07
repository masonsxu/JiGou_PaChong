import time
import requests
import json
import re
import pandas as pd
from selenium import webdriver
from bs4 import BeautifulSoup
from fake_useragent import UserAgent


# 判断是不是正确的json格式数据
def is_json(html):
    try:
        json.loads(html)
    except ValueError:
        return False
    return True


# 模拟谷歌浏览器
def get_html(url):
    ua = UserAgent(verify_ssl=False)  # 调用UserAgent库生成ua对象
    header = {
        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/web,image/png,*/*;q=0.8',
        'user-agent': ua.random,
    }
    try:
        r = requests.get(url, timeout=30, headers=header)
        r.raise_for_status()
        r.encoding = r.apparent_encoding  # 指定编码形式
        return r.text
    except:
        return "please inspect your url or setup"

URL_list = []
driver = webdriver.Chrome()
def connent_url(location):
    time.sleep(3)
    url = 'https://las.cnas.org.cn/LAS/publish/externalQueryPT.jsp'
    driver.get(url)

    orgAddress = driver.find_elements_by_id('orgAddress')
    orgAddress[0].send_keys(location)

    btn = driver.find_element_by_class_name('btn')
    btn.click()
    time.sleep(5)

    accept = False
    while not accept:
        try:
            pirlbutton1 = driver.find_element_by_xpath('//*[@id="pirlbutton1"]')
            print('Login')
            pirlbutton1.click()
            time.sleep(3)

            flag_str = driver.find_element_by_id('pirlAuthInterceptDiv_c').is_displayed()
            print(flag_str)

            if not flag_str:
                break

        except ValueError:
            accept = False
            time.sleep(3)

    flag_str = driver.find_element_by_xpath('/html/body/div[6]/div[1]/div[1]').is_displayed()
    if not flag_str:
        maxpage = int(driver.find_element_by_id('yui-pg0-0-totalPages-span').text)

        for page in range(maxpage):
            html = driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            tbody_all = soup.find_all('tbody')
            for url in tbody_all[1].find_all('a'):
                str1 = str(url)
                str2 = "https://las.cnas.org.cn" + str1[48:117]
                if len(str2) == 92:
                    URL_list.append(str2)

            if page <= maxpage-2:
                print(page)
                page_next = driver.find_element_by_xpath('/html/body/div[5]/table/tbody/tr/td[4]/span/img')
                page_next.click()
                time.sleep(3)
                accept = False
                while not accept:
                    try:
                        pirlbutton1 = driver.find_element_by_xpath('//*[@id="pirlbutton1"]')
                        pirlbutton1.click()
                        time.sleep(3)

                        flag_str = driver.find_element_by_id('pirlAuthInterceptDiv_c').is_displayed()

                        if not flag_str:
                            break

                    except UnicodeDecodeError:
                        accept = False
                        time.sleep(3)


# 解析目标网页的html
First_url_Json = []
Second_url_Json = []
Third_url_Json = []
Forth_url_Json = []
Datalist = []
def get_information_from_url(url):
    text = get_html(url)
    soup = BeautifulSoup(text, "html.parser")  # 解析text中的HTML

    first_url_Json = []
    second_url_Json = []
    third_url_Json = []
    for a_data in soup.find_all('a'):
        a_data = str(a_data)
        onclick_list = re.findall(".*onclick=(.*)'/LAS.*", a_data)
        onclick_str = ''
        for onclick in onclick_list:
            onclick_str = onclick_str + onclick
        onclick_str = onclick_str[1:-1]

        type_data1 = a_data[206:208]
        type_data2 = a_data[217:219]
        type_data3 = a_data[202:204]

        if type_data1 == "L1" or type_data2 == "L1":
            if onclick_str == "_showTop":
                id_list = re.findall(".*evaluateId=(.*)&amp;labType.*", a_data)
                id_str = ''
                for id in id_list:
                    id_str = id_str + id
                first_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishSignatory.action?&asstId=' + id_str
                second_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishPTAbilityQuery.action?' \
                             '&abilityType=L1&evaluateId=' + id_str
                if len(first_url_Json) == 0:
                    first_url_Json.append(first_url)
                if len(second_url_Json) == 0:
                    second_url_Json.append(second_url)

            elif onclick_str == "_showdown":
                id_list = re.findall(".*baseInfoId=(.*)&amp;labType.*", a_data)
                id_str = ''
                for id in id_list:
                    id_str = id_str + id
                first_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishSignatory.action?&baseinfoId=' + id_str
                second_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishPTAbilityQuery.action?' \
                             '&abilityType=L1&baseInfoId=' + id_str
                if len(first_url_Json) == 0:
                    first_url_Json.append(first_url)
                if len(second_url_Json) == 0:
                    second_url_Json.append(second_url)

        elif type_data1 == "L2" or type_data2 == "L2":
            if onclick_str == '_showTop':
                id_list = re.findall(".*evaluateId=(.*)&amp;labType.*", a_data)
                id_str = ''
                for id in id_list:
                    id_str = id_str + id
                first_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishKeyBranchY.action?&asstId=' + id_str
                second_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishSignatoryY.action?&evaluateId=' + id_str
                if len(first_url_Json) == 0:
                    first_url_Json.append(first_url)
                if len(second_url_Json) == 0:
                    second_url_Json.append(second_url)

            elif onclick_str == '_showdown':
                id_list = re.findall(".*baseInfoId=(.*)&amp;labType.*", a_data)
                id_str = ''
                for id in id_list:
                    id_str = id_str + id
                first_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishSignatory.action?&baseinfoId=' + id_str
                second_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishPTAbilityQuery.action?' \
                             '&abilityType=L1&baseInfoId=' + id_str
                third_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishPTAbilityQuery.action?' \
                            '&abilityType=L2&baseInfoId=' + id_str
                if len(first_url_Json) == 0:
                    first_url_Json.append(first_url)
                if len(second_url_Json) == 0:
                    second_url_Json.append(second_url)
                if len(third_url_Json) == 0:
                    third_url_Json.append(third_url)

        elif type_data3 == "L3":
            if onclick_str == '_showTop':
                id_list = re.findall(".*evaluateId=(.*)&amp;labType.*", a_data)
                id_str = ''
                for id in id_list:
                    id_str = id_str + id
                first_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishSignatoryY.action?&evaluateId=' + id_str
                second_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishPTAbilityYQuery.action?' \
                             '&abilityType=L3&evaluateId=' + id_str
                if len(first_url_Json) == 0:
                    first_url_Json.append(first_url)
                if len(second_url_Json) == 0:
                    second_url_Json.append(second_url)

            elif onclick_str == '_showdown':
                id_list = re.findall(".*baseInfoId=(.*)&amp;labType.*", a_data)
                id_str = ''
                for id in id_list:
                    id_str = id_str + id
                first_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishSignatory.action?&baseinfoId=' + id_str
                second_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishPTAbilityQuery.action?' \
                             '&abilityType=L3&baseInfoId=' + id_str
                if len(first_url_Json) == 0:
                    first_url_Json.append(first_url)
                if len(second_url_Json) == 0:
                    second_url_Json.append(second_url)


    data = []
    div_data = [div.get_text().strip().replace('\r', '').replace('\t', '').replace('\n', '').replace('\x1b', '')
                for div in soup.find_all('div', class_='T1')]

    for colomns_data in div_data:  # 遍历机构名称
        first_url_Json.append(colomns_data)
        second_url_Json.append(colomns_data)
        third_url_Json.append(colomns_data)
        data.append(colomns_data)

    for tr_data in soup.find_all('tr'):
        for span_data in tr_data.find_all('span', class_='clabel'):
            span_data = span_data.get_text().strip().replace('\x1b', '')  # 抽取表格内的数据
            data.append(span_data)

    if len(data) < 13:
        data_len = 13 - len(data)
        for num in range(data_len):
            data.append('')
    elif len(data) > 13:  # 删除非当前表格内的数据
        data_len = len(data) - 13
        for num in range(data_len):
            data.pop(-1)  # 从list尾部删除

    Datalist.append(data)
    if len(data) >= 9:
        span_str = str(data[10]+' - '+data[11]).replace('\r','').replace('\n','').replace('\t','')# 认证日期时间段
    else:
        span_str = ''
    first_url_Json.append(span_str)
    second_url_Json.append(span_str)
    third_url_Json.append(span_str)

    if len(first_url_Json[0]) > 100:
        First_url_Json.append(first_url_Json)
    if len(second_url_Json[0]) > 100:
        Second_url_Json.append(second_url_Json)
    if len(third_url_Json[0]) > 100:
        Third_url_Json.append(third_url_Json)

First_Data = []
def get_first_data(url):
    text = get_html(url[0])
    soup = BeautifulSoup(text, "html.parser")  # 解析text中的HTML
    html = soup.text
    flag_str = is_json(html)
    if flag_str is True:
        json_data = json.loads(html)
    else:
        json_data = {"data": []}

    infos = json_data['data']
    for info in infos:
        try:
            if len(url) == 3:
                zuzhi_name = str(url[1]).strip()
                date_str = str(url[2]).strip()
            else:
                zuzhi_name = str(url[2]).strip()
                date_str = str(url[3]).strip()
            data_list = []
            data_list.append(zuzhi_name)
            data_list.append(date_str)
            # 序号num
            try:
                num = info['num']
            except:
                num = ""
            if num is not None:
                data_list.append(num)
            else:
                data_list.append('')

            # 姓名nameCh
            try:
                nameCh = info['nameCh'].strip(

                ).replace('\r', '').replace('\n', '').replace('\t', '').replace('\x1b', '')

            except:
                nameCh = ''
            if nameCh is not None:
                data_list.append(nameCh)
            else:
                data_list.append('')

            # 授权签字领域authorizedFieldCh
            try:
                authorizedFieldCh = info['authorizedFieldCh'].strip().replace('\x1b', '')

            except:
                authorizedFieldCh = ""
            if authorizedFieldCh is not None:
                data_list.append(authorizedFieldCh)
            else:
                data_list.append('')

            # 说明note
            try:
                note = info['note'].strip().replace('\x1b', '')
            except:
                note = ""
            if note is None:
                data_list.append('')
            else:
                data_list.append(note)

            # 状态status
            try:
                status = info['status'].replace('\x1b', '')
            except:
                status = ""
            if status == '1':
                data_list.append("新增")
            else:
                data_list.append("有效")

            First_Data.append(data_list)

        except:
            sss = '本次爬取信息失败'
            print(sss)
            pass

Second_Data = []
def get_second_data(url):
    text = get_html(url[0])
    soup = BeautifulSoup(text, "html.parser")  # 解析text中的HTML
    html = soup.text
    flag_str = is_json(html)
    if flag_str is True:
        json_data = json.loads(html)
    else:
        json_data = {"data": []}

    infos = json_data['data']
    for info in infos:
        try:
            if len(url) == 3:
                zuzhi_name = str(url[1]).strip()
                date_str = str(url[2]).strip()
            else:
                zuzhi_name = str(url[2]).strip()
                date_str = str(url[3]).strip()
            data_list = []
            data_list.append(zuzhi_name)
            data_list.append(date_str)
            # 序号num
            try:
                num = info['num']
            except:
                num = ""
            if num is not None:
                data_list.append(num)
            else:
                data_list.append('')

            # 样品名称nameCh
            try:
                nameCh = info['nameCh'].strip().replace('\x1b', '')

            except:
                nameCh = ''
            if nameCh is not None:
                data_list.append(nameCh)
            else:
                data_list.append('')

            # 项目参数序号detNum
            try:
                detNum = info['detNum']
            except:
                detNum = ""
            if detNum is not None:
                data_list.append(detNum)
            else:
                data_list.append('')

            # 项目名称itemCh
            try:
                itemCh = info['itemCh'].strip().replace('\x1b', '')

            except:
                itemCh = ""
            if itemCh is not None:
                data_list.append(itemCh)
            else:
                data_list.append('')

            # 说明detNoteCh
            try:
                detNoteCh = info['detNoteCh'].strip(

                ).replace('\r', '').replace('\n', '').replace('\t', '').replace('\x1b', '')
            except:
                detNoteCh = ''
            if detNoteCh is None:
                data_list.append('')
            data_list.append(detNoteCh)

            # 状态status
            try:
                status = info['status'].replace('\x1b', '')
            except:
                status = ""
            if status == '1':
                data_list.append("新增")
            else:
                data_list.append("有效")

            Second_Data.append(data_list)

        except:
            sss = '本次爬取信息失败'
            print(sss)
            pass

Third_Data = []
def get_third_data(url):
    text = get_html(url[0])
    soup = BeautifulSoup(text, "html.parser")  # 解析text中的HTML
    html = soup.text
    flag_str = is_json(html)
    if flag_str is True:
        json_data = json.loads(html)
    else:
        json_data = {"data": []}

    infos = json_data['data']
    for info in infos:
        try:
            if len(url) == 3:
                zuzhi_name = str(url[1]).strip()
                date_str = str(url[2]).strip()
            else:
                zuzhi_name = str(url[2]).strip()
                date_str = str(url[3]).strip()
            data_list = []
            data_list.append(zuzhi_name)
            data_list.append(date_str)

            # 序号num
            try:
                num = info['num']
            except:
                num = ""
            if num is not None:
                data_list.append(str(num))
            else:
                data_list.append('')

            # 样品名称nameCh
            try:
                nameCh = info['nameCh'].strip(

                ).replace('\r', '').replace('\n', '').replace('\t', '').replace('\x1b', '')
            except:
                nameCh = ''
            if nameCh is not None:
                data_list.append(nameCh)
            else:
                data_list.append('')

            # 检验项目序号detNum
            try:
                detNum = info['detNum']
            except:
                detNum = " "
            if detNum is not None:
                data_list.append(str(detNum))
            else:
                data_list.append('')

            # 检验项目名称itemCh
            try:
                itemCh = info['itemCh'].strip(

                ).replace('\r', '').replace('\n', '').replace('\t', '').replace('\x1b', '')
            except:
                itemCh = ""
            if itemCh is not None:
                data_list.append(itemCh)
            else:
                data_list.append('')

            # 说明detNoteCh
            try:
                detNoteCh = info['detNoteCh'].strip(

                ).replace('\r', '').replace('\n', '').replace('\t', '').replace('\x1b', '')

            except:
                detNoteCh = ''
            if detNoteCh is None:
                data_list.append('')
            data_list.append(detNoteCh)

            # 状态status
            try:
                status = info['status'].strip().replace('\x1b', '')
            except:
                status = ''
            if status == '1':
                data_list.append("新增")
            else:
                data_list.append("有效")

            Third_Data.append(data_list)

        except:
            sss = '本次爬取信息失败'
            print(sss)
            pass

if __name__ == '__main__':
    df = pd.read_excel('E:\\Pycharm_xjf\\JiGou_PaChong\\Nengli_Data\\机构地址.xlsx', sheet_name='Sheet1')
    data = df['机构地址']

    i = 1
    for key in data:
        connent_url(str(key))
        i = i + 1
    driver.close()
    driver.quit()

    df_Sheet1 = pd.DataFrame(URL_list, columns=['url'])
    df_Sheet1.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Nengli_Data\\Nengli_url1.csv')
    i = 1
    for url in URL_list:
        get_information_from_url(url)
        print(i)
        i = i + 1
    df_Sheet2 = pd.DataFrame(Datalist, columns=['机构名称', '注册编号', '报告/证书允许使用认可标识的其他名称', '联系人',
                                                '联系电话', '邮政编码', '传真号码', '网站地址', '电子邮箱', '单位地址',
                                                '认证开始时间', '认证结束时间', '认可依据'])
    print("Sheet1 - Successful!!!")
    df_Sheet2.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Nengli_Data\\Nengli_Data1.csv')

    df_Sheet3 = pd.DataFrame(First_url_Json)
    df_Sheet3.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Nengli_Data\\Nengli_url2.csv')
    i = 1
    for url in First_url_Json:
        get_first_data(url)
        print(i)
        i = i+1
    df_Sheet4 = pd.DataFrame(First_Data, columns=['机构名称', '周期', '序号', '姓名', '授权签字领域', '说明',  '状态'])
    print('Sheet2 - Successful!!!')
    df_Sheet4.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Nengli_Data\\Nengli_Data2.csv')

    df_Sheet5 = pd.DataFrame(Second_url_Json)
    df_Sheet5.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Nengli_Data\\Nengli_url3.csv')
    i = 1
    for url in Second_url_Json:
        get_second_data(url)
        print(i)
        i = i+1
    df_Sheet6 = pd.DataFrame(Second_Data, columns=['机构名称', '序号', '周期', '样品名称', '项目序号', '项目名称', '说明',
                                                   '状态'])
    print('Sheet3 - Successful!!!')
    df_Sheet6.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Nengli_Data\\Nengli_Data3.csv')

    df_Sheet7 = pd.DataFrame(Third_url_Json)
    df_Sheet7.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Nengli_Data\\Nengli_url4.csv')
    i = 1
    for url in Third_url_Json:
        get_third_data(url)
        print(i)
        i = i + 1
    df_Sheet8 = pd.DataFrame(Third_Data, columns=['机构名称', '周期', '序号', '样品名称', '项目序号', '检验项目名称',
                                                  '说明', '状态'])
    print('Sheet4 - Successful!!!')
    df_Sheet8.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Nengli_Data\\Nengli_Data4.csv')

    outfile = 'E:\\Pycharm_xjf\\JiGou_PaChong\\Nengli_Data\\Nengli_Data.xlsx'
    writer = pd.ExcelWriter(outfile)
    df_Sheet2.to_excel(excel_writer=writer, sheet_name='机构信息', index=None)
    df_Sheet4.to_excel(excel_writer=writer, sheet_name='认可的授权签字人及领域', index=None)
    df_Sheet6.to_excel(excel_writer=writer, sheet_name='认可的检测领域范围', index=None)
    df_Sheet8.to_excel(excel_writer=writer, sheet_name='认可的校准能力范围', index=None)
    writer.save()
    print('save - Successful!!!')
    writer.close()
    print('close - Successful!!!')
