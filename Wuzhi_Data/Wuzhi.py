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
    url = 'https://las.cnas.org.cn/LAS/publish/externalQueryRM.jsp'
    driver.get(url)

    orgAddress = driver.find_elements_by_id('orgAddress')
    orgAddress[0].send_keys(location)

    btn = driver.find_element_by_class_name('btn')
    btn.click()
    time.sleep(5)

    accept = False
    while not accept:
        try:
            pirlbutton1 = driver.find_element_by_xpath('//*[@id="pirlbutton1"]')  # 点击确认
            print('Login')
            pirlbutton1.click()
            time.sleep(3)
            # 判断查询是否成功
            flag_str = driver.find_element_by_id('pirlAuthInterceptDiv_c').is_displayed()
            print(flag_str)

            if not flag_str:
                break

        except ValueError:
            accept = False
            time.sleep(3)

    flag_str = driver.find_element_by_xpath('/html/body/div[6]/div[1]/div[1]').is_displayed()  # 判断当前页数
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
                page_next = driver.find_element_by_xpath('/html/body/div[5]/table/tbody/tr/td[4]/span/img')  # 定位换页
                page_next.click()
                time.sleep(3)
                accept = False
                while not accept:
                    try:
                        pirlbutton1 = driver.find_element_by_xpath('//*[@id="pirlbutton1"]')  # 点击确认
                        pirlbutton1.click()
                        time.sleep(3)
                        # 判断搜索信息是否成功
                        flag_str = driver.find_element_by_id('pirlAuthInterceptDiv_c').is_displayed()

                        if not flag_str:
                            break

                    except UnicodeDecodeError:
                        accept = False
                        time.sleep(3)


# 解析目标网页的html
First_url_Json = []
Second_url_Json = []
Datalist = []
def get_information_from_url(url):
    text = get_html(url)
    soup = BeautifulSoup(text, "html.parser")  # 解析text中的HTML

    first_url_Json = []
    for a_data in soup.find_all('a'):
        a_data = str(a_data)
        onclick_list = re.findall(".*onclick=(.*)'/LAS.*", a_data)  # 截取字符串，判断认证类型（正式/非正式）
        onclick_str = ''
        for onclick in onclick_list:
            onclick_str = onclick_str + onclick
        onclick_str = onclick_str[1:-1]

        if onclick_str == "_showTop":
            id_list = re.findall(".*evaluateId=(.*)&amp;labType.*", a_data)  # 截取evaluateId的ID
            id_str = ''
            for id in id_list:
                id_str = id_str + id
            first_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishRMAbilityQuery.action?&evaluateId=' + id_str
            if len(first_url_Json) == 0:
                first_url_Json.append(first_url)

        elif onclick_str == "_showdown":
            id_list = re.findall(".*baseInfoId=(.*)&amp;labType.*", a_data)  # 截取baseInfoId的ID
            id_str = ''
            for id in id_list:
                id_str = id_str + id
            first_url = 'https://las.cnas.org.cn/LAS/publish/queryPublishRMAbilityQuery.action?&baseInfoId=' + id_str
            if len(first_url_Json) == 0:
                first_url_Json.append(first_url)


    data = []
    div_data = [div.get_text().strip().replace('\r', '').replace('\t', '').replace('\n', '').replace('\x1b', '')
                for div in soup.find_all('div', class_='T1')]

    for colomns_data in div_data:  # 遍历机构名称
        first_url_Json.append(colomns_data)
        data.append(colomns_data)

    for tr_data in soup.find_all('tr'):
        for span_data in tr_data.find_all('span', class_='clabel'):
            span_data = span_data.get_text().strip(

            ).replace('\r', '').replace('\n', '').replace('\t', '').replace('\x1b', '')  # 抽取表格内的数据
            data.append(span_data)

    if len(data) < 12:
        data_len = 12 - len(data)
        for num in range(data_len):
            data.append('')
    elif len(data) > 12:  # 删除非当前表格内的数据
        data_len = len(data) - 12
        for num in range(data_len):
            data.pop(-1)  # 从list尾部删除

    Datalist.append(data)
    if len(data) >= 9:
        span_str = str(data[10]+' - '+data[11]).replace('\r', '').replace('\n', '').replace('\t', '')  # 认证日期时间段
    else:
        span_str = ''
    first_url_Json.append(span_str)

    if len(first_url_Json[0]) > 100:
        First_url_Json.append(first_url_Json)

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
            # 序号
            try:
                num = info['num']
            except:
                num = ""
            if num is not None:
                data_list.append(num)
            else:
                data_list.append('')

            # 名称name
            try:
                name = info['name'].strip().replace('\x1b', '')

            except:
                name = ''
            if name is not None:
                data_list.append(name)
            else:
                data_list.append('')

            # 成分特性序号elemNum
            try:
                elemNum = info['elemNum'].strip().replace('\x1b', '')

            except:
                elemNum = ""
            if elemNum is not None:
                data_list.append(elemNum)
            else:
                data_list.append('')

            # 特性范围frangeVal
            try:
                frangeVal = info['frangeVal'].strip(

                ).replace('\r', '').replace('\n', '').replace('\t', '').replace('\x1b', '')
            except:
                frangeVal = ''
            if frangeVal is None:
                data_list.append('')
            data_list.append(frangeVal)

            # 不确定度nrangeVal
            try:
                nrangeVal = info['nrangeVal'].strip(

                ).replace('\r', '').replace('\n', '').replace('\t', '').replace('\x1b', '')
            except:
                nrangeVal = ''
            if nrangeVal is None:
                data_list.append('')
            data_list.append(nrangeVal)

            # 定值方式valueType
            try:
                valueType = info['valueType'].strip(

                ).replace('\r', '').replace('\n', '').replace('\t', '').replace('\x1b', '')
            except:
                valueType = ''
            if valueType is None:
                data_list.append('')
            data_list.append(valueType)

            # 测试技术testSkill
            try:
                testSkill = info['testSkill'].strip().replace('\r', '').replace('\n', '').replace('\t', '').replace('\x1b', '')
            except:
                testSkill = ''
            if testSkill is None:
                data_list.append('')
            data_list.append(testSkill)

            # 说明explain
            try:
                explain = info['explain'].strip().replace('\r', '').replace('\n', '').replace('\t', '').replace('\x1b', '')
            except:
                explain = ''
            if explain is None:
                data_list.append('')
            data_list.append(explain)

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


if __name__ == '__main__':
    df = pd.read_excel('E:\\Pycharm_xjf\\JiGou_PaChong\\Wuzhi_Data\\机构地址.xlsx', sheet_name='Sheet1')
    data = df['机构地址']

    i = 1
    for key in data:
        connent_url(str(key))
        i = i + 1
    driver.close()
    driver.quit()

    df_Sheet1 = pd.DataFrame(URL_list, columns=['url'])
    df_Sheet1.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Wuzhi_Data\\Wuzhi_url1.csv')
    i = 1
    for url in URL_list:
        get_information_from_url(url)
        print(i)
        i = i + 1
    df_Sheet2 = pd.DataFrame(Datalist, columns=['机构名称', '注册编号', '报告/证书允许使用认可标识的其他名称', '联系人',
                                                '联系电话', '邮政编码', '传真号码', '网站地址', '电子邮箱', '单位地址',
                                                '认证开始时间', '认证结束时间'])
    print("Sheet1 - Successful!!!")
    df_Sheet2.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Wuzhi_Data\\Wuzhi_Data1.csv')

    df_Sheet3 = pd.DataFrame(First_url_Json)
    df_Sheet3.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Wuzhi_Data\\Wuzhi_url2.csv')
    i = 1
    for url in First_url_Json:
        get_first_data(url)
        print(i)
        i = i+1
    df_Sheet4 = pd.DataFrame(First_Data, columns=['机构名称', '周期', '序号', '名称', '成分特性名称', '特性范围', '不确定度',
                                                  '定值方式', '测试技术', '说明', '状态'])
    print('Sheet2 - Successful!!!')
    df_Sheet4.to_csv('E:\\Pycharm_xjf\\JiGou_PaChong\\Wuzhi_Data\\Wuzhi_Data2.csv')

    outfile = 'E:\\Pycharm_xjf\\JiGou_PaChong\\Wuzhi_Data\\Wuzhi_Data.xlsx'
    writer = pd.ExcelWriter(outfile)
    df_Sheet2.to_excel(excel_writer=writer, sheet_name='机构信息', index=None)
    df_Sheet4.to_excel(excel_writer=writer, sheet_name='认可的授权签字人及领域', index=None)
    writer.save()
    print('save - Successful!!!')
    writer.close()
    print('close - Successful!!!')
