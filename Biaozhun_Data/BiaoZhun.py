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
    ua = UserAgent()  # 调用UserAgent库生成ua对象
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
    url = 'https://las.cnas.org.cn/LAS/publish/externalQueryL1.jsp'
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

    flag_str = driver.find_element_by_xpath('/html/body/div[7]/div[1]/div[1]').is_displayed()
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
                page_next = driver.find_element_by_xpath('/html/body/div[6]/table/tbody/tr/td[4]/a/img')
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
    forth_url_Json = []
    for a_data in soup.find_all('a'):
        a_data = str(a_data)
        onclick_list = re.findall(".*onclick=(.*)'/LAS.*", a_data)
        onclick_str = ''
        for onclick in onclick_list:
            onclick_str = onclick_str + onclick
        onclick_str = onclick_str[1:-1]

        type_data1 = a_data[205:207]
        type_data2 = a_data[201:203]
        if type_data1 == "L1" or type_data2 == "L1":
            if onclick_str == "_showTop":
                id_list = re.findall(".*evaluateId=(.*)&amp;labType.*", a_data)
                id_str = ''
                for id in id_list:
                    id_str = id_str + id
                first_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishKeyBranchY.action?&asstId=' + id_str
                second_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishSignatoryY.action?&evaluateId=' + id_str
                third_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishLCheckObjY.action?&evaluateId=' \
                            + id_str + '&type=L1'
                if len(first_url_Json) == 0:
                    first_url_Json.append(first_url)
                if len(second_url_Json) == 0:
                    second_url_Json.append(second_url)
                if len(third_url_Json) == 0:
                    third_url_Json.append(third_url)

            elif onclick_str == "_showdown":
                id_list = re.findall(".*baseInfoId=(.*)&amp;labType.*", a_data)
                id_str = ''
                for id in id_list:
                    id_str = id_str + id
                first_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishKeyBranch.action?&asstId=' + id_str
                second_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishSignatory.action?&baseinfoId=' + id_str
                third_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishLCheckObj.action?&baseinfoId=' \
                            + id_str + '&type=L1'
                if len(first_url_Json) == 0:
                    first_url_Json.append(first_url)
                if len(second_url_Json) == 0:
                    second_url_Json.append(second_url)
                if len(third_url_Json) == 0:
                    third_url_Json.append(third_url)

        elif type_data1 == "L2" or type_data2 == "L2":
            if onclick_str == '_showTop' and (len(a_data and "L1") > 0):
                id_list = re.findall(".*evaluateId=(.*)&amp;labType.*", a_data)
                id_str = ''
                for id in id_list:
                    id_str = id_str + id
                first_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishKeyBranchY.action?&asstId=' + id_str
                second_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishSignatoryY.action?&evaluateId=' + id_str
                third_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishLCheckObjY.action?&evaluateId=' \
                            + id_str + '&type=L1'
                forth_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishLCailObjY.action?&evaluateId=' \
                            + id_str + '&type=L2'
                if len(first_url_Json) == 0:
                    first_url_Json.append(first_url)
                if len(second_url_Json) == 0:
                    second_url_Json.append(second_url)
                if len(third_url_Json) == 0:
                    third_url_Json.append(third_url)
                if len(forth_url_Json) == 0:
                    forth_url_Json.append(forth_url)

            elif onclick_str == '_showdown':
                id_list = re.findall(".*baseInfoId=(.*)&amp;labType.*", a_data)
                id_str = ''
                for id in id_list:
                    id_str = id_str + id
                first_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishKeyBranch.action?&asstId=' + id_str
                second_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishSignatory.action?&baseinfoId=' + id_str
                third_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishLCailObj.action?&baseinfoId=' \
                            + id_str + '&type=L2'
                if len(first_url_Json) == 0:
                    first_url_Json.append(first_url)
                if len(second_url_Json) == 0:
                    second_url_Json.append(second_url)
                if len(third_url_Json) == 0:
                    third_url_Json.append(third_url)

    data = []
    div_data = [div.get_text().strip().replace('\r', '').replace('\t', '').replace('\n', '').replace('\x1b', '')
                for div in soup.find_all('div', class_='T1')]

    for colomns_data in div_data:  # 遍历机构名称
        first_url_Json.append(colomns_data)
        second_url_Json.append(colomns_data)
        third_url_Json.append(colomns_data)
        forth_url_Json.append(colomns_data)
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
        span_str = str(data[10]+' - '+data[11]).replace('\r', '').replace('\n', '').replace('\t', '')  # 认证日期时间段
    else:
        span_str = ''
    first_url_Json.append(span_str)
    second_url_Json.append(span_str)
    third_url_Json.append(span_str)
    forth_url_Json.append(span_str)

    if len(first_url_Json[0]) > 100:
        First_url_Json.append(first_url_Json)
    if len(second_url_Json[0]) > 100:
        Second_url_Json.append(second_url_Json)
    if len(third_url_Json[0]) > 100:
        Third_url_Json.append(third_url_Json)
    if len(forth_url_Json[0]) > 100:
        Forth_url_Json.append(forth_url_Json)

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
            # 地址代码
            try:
                keyNum = chr(info['keyNum'] + 64)

            except:
                keyNum = ""
            if keyNum is not None:
                data_list.append(keyNum)
            else:
                data_list.append('')

            # 地址
            try:
                addCn = info['addCn'].strip().replace('\r', '').replace('\n', '').replace('\t', '').replace('\x1b', '')

            except:
                addCn = ''
            if addCn is not None:
                data_list.append(addCn)
            else:
                data_list.append('')

            # 邮编
            try:
                postCode = info['postCode'].strip().replace('\x1b', '')

            except:
                postCode = ""
            if postCode is not None:
                data_list.append(postCode)
            else:
                data_list.append('')

            # 设施特点labFeatureJson
            try:
                labFeatureJson = str(info['labFeatureJson'].replace('\x1b', ''))

            except:
                labFeatureJson = ""
            labFeatureJson = json.loads(labFeatureJson)
            feature_str = ""
            for feature_json in labFeatureJson:
                feature_str = feature_str + feature_json['feature'] + ','
            feature_str = feature_str[:-1]
            data_list.append(feature_str)

            # 主要活动mainactivity
            try:
                mainactivity = info['mainactivity'].strip().replace('\x1b', '')

            except:
                mainactivity = ""
            if mainactivity is not None:
                data_list.append(mainactivity)
            else:
                data_list.append('')

            #说明remark
            try:
                remark = info['remark'].strip().replace('\x1b','')
            except:
                remark = ""
            if remark is None:
                data_list.append('')
            else:
                data_list.append(remark)

            # 是否推荐primaryRecommend
            try:
                primaryRecommend = info['primaryRecommend'].replace('\x1b', '')
            except:
                primaryRecommend = ""
            if primaryRecommend == '1':
                data_list.append('是')
            else:
                data_list.append('')

            # 状态bstatus
            try:
                isModify = info['isModify'].replace('\x1b', '')
            except:
                isModify = ""
            if isModify == '1':
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
            # 序号
            try:
                num = info['num']

            except:
                num = ""
            if num is not None:
                data_list.append(num)
            else:
                data_list.append('')

            # 姓名
            try:
                nameCh = info['nameCh'].strip().replace('\x1b', '')

            except:
                nameCh = ''
            if nameCh is not None:
                data_list.append(nameCh)
            else:
                data_list.append('')

            # 授权签字领域
            try:
                authorizedFieldCh = info['authorizedFieldCh'].strip().replace('\x1b', '')

            except:
                authorizedFieldCh = ""
            if authorizedFieldCh != None:
                data_list.append(authorizedFieldCh)
            else:
                data_list.append('')

            # 说明
            try:
                note = info['note'].strip().replace('\r', '').replace('\n', '').replace('\t', '').replace('\x1b', '')
            except:
                note = ''
            if note is None:
                data_list.append('')
            data_list.append(note)

            #是否推荐recommend
            try:
                recommend = info['recommend'].replace('\x1b', '')
            except:
                recommend = ""
            if recommend == '1':
                data_list.append('是')
            else:
                data_list.append('')

            # 状态bstatus
            try:
                isModify = info['isModify'].replace('\x1b', '')
            except:
                isModify = ""
            if isModify == '1':
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
            # 分组名称
            try:
                typeName = info['typeName']

            except:
                typeName = ""
            if typeName is None:
                typeName = '未分组'
            typeName.strip()
            data_list.append(typeName)

            # 检验对象序号
            try:
                num = info['num']

            except:
                num = ""
            if num is not None:
                data_list.append(str(num))
            else:
                data_list.append('')

            #检验对象
            try:
                objCh = info['objCh'].strip().replace('\r', '').replace('\n',
                                                                        '').replace('\t', '').replace('\x1b', '')

            except:
                objCh = ''
            if objCh is not None:
                data_list.append(objCh)
            else:
                data_list.append('')

            # 检验项目名称序号
            try:
                paramNum = info['paramNum']

            except:
                paramNum = " "
            if paramNum is not None:
                data_list.append(str(paramNum))
            else:
                data_list.append('')

            # 检验项目名称
            try:
                paramCh = info['paramCh'].strip().replace('\r', '').replace('\n',
                                                                            '').replace('\t', '').replace('\x1b', '')

            except:
                paramCh = ""
            if paramCh is not None:
                data_list.append(paramCh)
            else:
                data_list.append('')

            # 依据的检测标准
            try:
                stdAllDesc = info['stdAllDesc'].strip().replace('\x1b', '')

            except:
                stdAllDesc = ""
            if stdAllDesc is not None:
                data_list.append(stdAllDesc)
            else:
                data_list.append('')

            #说明
            try:
                limitCh = info['limitCh'].strip().replace('\r', '').replace('\n', '').replace('\t',
                                                                                              '').replace('\x1b', '')

            except:
                limitCh = ''
            if limitCh is None:
                data_list.append('')
            data_list.append(limitCh)

            # 状态stdStatus
            try:
                stdStatus = info['stdStatus'].strip().replace('\x1b', '')
            except:
                stdStatus = ''
            if stdStatus == '0':
                stdStatus = "有效"
            else:
                stdStatus = "新增"
            data_list.append(stdStatus)


            Third_Data.append(data_list)

        except:
            sss = '本次爬取信息失败'
            print(sss)
            pass

Forth_Data = []
def get_forth_data(url):
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
            # 分组名称
            try:
                typeName = info['typeName']
            except:
                typeName = ""
            if typeName is None:
                typeName = '未分组'
            typeName.strip()
            data_list.append(typeName)

            # 序号objNum
            try:
                objNum = info['objNum']

            except:
                objNum = ""
            if objNum is not None:
                data_list.append(str(objNum))
            else:
                data_list.append('')

            # 测量仪器名称objCh
            try:
                objCh = info['objCh'].strip().replace('\r', '').replace('\n',
                                                                        '').replace('\t', '').replace('\x1b', '')

            except:
                objCh = ''
            if objCh is not None:
                data_list.append(objCh)
            else:
                data_list.append('')

            # 被测量项目名称paramCh
            try:
                paramCh = info['paramCh'].strip().replace('\x1b','')

            except:
                paramCh = ""
            if paramCh is not None:
                data_list.append(paramCh)
            else:
                data_list.append('')

            # 校准范围standardCodeStr
            try:
                standardCodeStr = info['standardCodeStr'].strip().replace('\x1b', '')
            except:
                standardCodeStr = ""
            if standardCodeStr is not None:
                data_list.append(standardCodeStr)
            else:
                data_list.append('')

            # 测量范围testCh
            try:
                testCh = info['testCh'].strip().replace('\x1b', '')
            except:
                testCh = ''
            if testCh is not None:
                data_list.append(testCh)
            else:
                data_list.append('')

            # 扩展不确度kvalueTag kvalueCh
            try:
                kvalueTag = info['kvalueTag'].strip().replace('\x1b', '')
                kvalueCh = info['kvalueCh'].strip().replace('\x1b', '')
            except:
                kvalueTag = ''
                kvalueCh = ''
            if kvalueTag is not None or kvalueCh is not None:
                data_list.append(kvalueTag + "=" + kvalueCh)
            else:
                data_list.append('')

            # 说明limitCh
            try:
                limitCh = info['limitCh'].strip().replace('\r', '').replace('\n',
                                                                            '').replace('\t', '').replace('\x1b', '')
            except:
                limitCh = ''
            if limitCh is None:
                data_list.append('')
            data_list.append(limitCh)

            # 状态status
            try:
                status = info['status'].strip().replace('\x1b','')
            except:
                status = ''
            if status == '0':
                status = "有效"
            else:
                status = "新增"
            data_list.append(status)

            Forth_Data.append(data_list)



        except:
            sss = '本次爬取信息失败'
            print(sss)
            pass

if __name__ == '__main__':
    df = pd.read_excel('E:\\Pycharm_xjf\\JiGou_PaChong\\Biaozhun_Data\\机构地址.xlsx', sheet_name='Sheet1')
    data = df['机构地址']

    i = 1
    for key in data:
        connent_url(str(key))
        i = i + 1
    driver.close()
    driver.quit()

    df_Sheet1 = pd.DataFrame(URL_list, columns=['url'])
    df_Sheet1.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Biaozhun_Data\\Biaozhun_url1.csv')
    i = 1
    for url in URL_list:
        get_information_from_url(url)
        print(i)
        i = i + 1
    df_Sheet2 = pd.DataFrame(Datalist, columns=['机构名称', '注册编号', '报告/证书允许使用认可标识的其他名称', '联系人',
                                                '联系电话', '邮政编码', '传真号码', '网站地址', '电子邮箱', '单位地址',
                                                '认证开始时间', '认证结束时间', '暂停项目/参数'])
    print("Sheet1 - Successful!!!")
    df_Sheet2.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Biaozhun_Data\\Biaozhun_Data1.csv')

    df_Sheet3 = pd.DataFrame(First_url_Json)
    df_Sheet3.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Biaozhun_Data\\Biaozhun_url2.csv')
    i = 1
    for url in First_url_Json:
        get_first_data(url)
        print(i)
        i = i+1
    df_Sheet4 = pd.DataFrame(First_Data, columns=['机构名称', '周期', '地址代码', '地址', '邮编', '设施特点', '主要活动',
                                                  '说明', '是否推荐', '状态'])
    print('Sheet2 - Successful!!!')
    df_Sheet4.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Biaozhun_Data\\Biaozhun_Data2.csv')

    df_Sheet5 = pd.DataFrame(Second_url_Json)
    df_Sheet5.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Biaozhun_Data\\Biaozhun_url3.csv')
    i = 1
    for url in Second_url_Json:
        get_second_data(url)
        print(i)
        i = i+1
    df_Sheet6 = pd.DataFrame(Second_Data, columns=['机构名称', '周期', '序号', '姓名', '授权签字领域', '说明',
                                                   '是否推荐', '状态'])
    print('Sheet3 - Successful!!!')
    df_Sheet6.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Biaozhun_Data\\Biaozhun_Data3.csv')

    df_Sheet7 = pd.DataFrame(Third_url_Json)
    df_Sheet7.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Biaozhun_Data\\Biaozhun_url4.csv')
    i = 1
    for url in Third_url_Json:
        get_third_data(url)
        print(i)
        i = i + 1
    df_Sheet8 = pd.DataFrame(Third_Data, columns=['机构名称', '周期', '分组名称', '检验对象序号', '检验对象',
                                                  '检验对象名称序号', '检验项目名称', '依据检验标准', '说明', '状态'])
    print('Sheet4 - Successful!!!')
    df_Sheet8.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Biaozhun_Data\\Biaozhun_Data4.csv')

    df_Sheet9 = pd.DataFrame(Forth_url_Json)
    df_Sheet9.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Biaozhun_Data\\Biaozhun_url5.csv')
    i = 1
    for url in Forth_url_Json:
        get_forth_data(url)
        print(i)
        i = i + 1
    df_Sheet10 = pd.DataFrame(Forth_Data, columns=['机构名称', '周期', '分组名称', '检验对象序号', '测量仪器名称',
                                                   '被测量项目名称', '校准范围', '测量范围', '扩展不确度', '说明', '状态'])
    print('Sheet5 - Successful!!!')
    df_Sheet10.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Biaozhun_Data\\Biaozhun_Data5.csv')

    outfile = 'E:\\Pycharm_xjf\\JiGou_PaChong\\Biaozhun_Data\\Biaozhun_Data.xlsx'
    writer = pd.ExcelWriter(outfile)
    df_Sheet2.to_excel(excel_writer=writer, sheet_name='机构信息', index=None)
    df_Sheet4.to_excel(excel_writer=writer, sheet_name='认可的实验室关键场所', index=None)
    df_Sheet6.to_excel(excel_writer=writer, sheet_name='认可的授权签字人及领域', index=None)
    df_Sheet8.to_excel(excel_writer=writer, sheet_name='认可的能力检测范围', index=None)
    df_Sheet10.to_excel(excel_writer=writer, sheet_name='认可的测量和校准范围', index=None)
    writer.save()
    print('save - Successful!!!')
    writer.close()
    print('close - Successful!!!')
