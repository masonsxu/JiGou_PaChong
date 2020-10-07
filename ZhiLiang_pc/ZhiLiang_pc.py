import requests
import pandas as pd
from bs4 import BeautifulSoup
from fake_useragent import UserAgent

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

# 分析当前页面信息，按照深层分析中的列表长度进行分类，2，3，4，5，2/4，将url信息分别存储下来，保存当前页面信息的全部数据
Url_List = []
First_Url_List_2 = []
First_Url_List_4 = []
First_Url_List_5 = []
First_Url_List_2_4 = []
Second_Url_List = []
Third_Url_List = []
Name_Data_List = []
def get_url_data(url_str):
    html = get_html(url_str)
    soup = BeautifulSoup(html, 'html.parser')
    table_view = soup.find_all('table', id='view')
    # print(table_view)
    tbody_tr = table_view[0].find_all('tbody')
    # print(tbody_tr)
    for cell_tr in tbody_tr[0].find_all('tr'):
        # print(cell_tr)
        name_data_list = []
        first_data_2 = []
        first_data_4 = []
        first_data_5 = []
        first_data_2_4 = []
        second_data = []
        third_data = []
        for cell_td in cell_tr.find_all('td'):
            # print(len(cell_td))
            if len(cell_td) > 1:
                a_url_str = 'http://cb.cnas.org.cn/cnas_published/servlet/NetShowServlet?actionType=showOrgDetail'
                a_first_url_str = 'http://cb.cnas.org.cn/cnas_published/servlet/NetShowServlet?actionType=showOrgScope'
                a_second_url_str = 'http://cb.cnas.org.cn/cnas_published/servlet/NetShowServlet?actionType=showOrgLimitCode'
                a_third_url_str = 'http://cb.cnas.org.cn/cnas_published/servlet/NetShowServlet?actionType=showSuborg'
                for cell_a in cell_td.find_all('a'):
                    onclick_str = str(cell_a.get('onclick'))
                    orgId = '&orgId=' + onclick_str[12:22]
                    field = '&field=' + onclick_str[25:-2]
                    path_url = a_url_str + orgId + field
                    Url_List.append(path_url)
                    path_url_first = a_first_url_str + orgId +field
                    path_url_second = a_second_url_str + orgId +field
                    second_data.append(path_url_second)
                    path_url_third = a_third_url_str + orgId + field
                    third_data.append(path_url_third)
                    if (field == '&field=HACCP') or (field == '&field=GMP') or (field == '&field=FSMS'):
                        first_data_2.append(path_url_first)
                    elif (field == '&field=OMS') or (field == '&field=EMS') or (field == '&field=OHSMS')\
                            or (field == '&field=ISMS'):
                        first_data_4.append(path_url_first)
                    elif (field == '&field=P') or (field == '&field=SC') or (field == '&field=LC') or (field == '&field=EW'):
                        first_data_5.append(path_url_first)
                    elif (field == '&field=GAP') or (field == '&field=OP') or (field == '&field=FR')\
                        or (field == '&field=SP') or (field == '&field=EnMS') or (field == '&field=TL')\
                            or (field == '&field=R') or (field == '&field=EC9000') or (field == '&field=ITSMS'):
                        first_data_2_4.append(path_url_first)
            cell_td_str = str(cell_td.text).strip()
            # print(cell_td_str)
            name_data_list.append(cell_td_str)

        Name_Data_List.append(name_data_list)

        name = name_data_list[0]
        second_data.append(name)
        third_data.append(name)
        date = name_data_list[3] + ' - ' + name_data_list[2]
        second_data.append(date)
        third_data.append(date)
        Second_Url_List.append(second_data)
        Third_Url_List.append(third_data)
        if len(first_data_2) != 0:
            first_data_2.append(name)
            first_data_2.append(date)
            First_Url_List_2.append(first_data_2)
        elif len(first_data_4) != 0:
            first_data_4.append(name)
            first_data_4.append(date)
            First_Url_List_4.append(first_data_4)
        elif len(first_data_5) != 0:
            first_data_5.append(name)
            first_data_5.append(date)
            First_Url_List_5.append(first_data_5)
        elif len(first_data_2_4) != 0:
            first_data_2_4.append(name)
            first_data_2_4.append(date)
            First_Url_List_2_4.append(first_data_2_4)
# 定义一个又返回列表的方法，返回当前的页面的信息，方便后续信息的整理分析
Information_Data_List = []
def get_next_data(url_str):
    html = get_html(url_str)
    soup = BeautifulSoup(html, 'html.parser')
    table_table1 = soup.find_all('table', class_='table1')
    information_data_list = []
    for cell_tr in table_table1[0].find_all('tr'):
        for cell_td in cell_tr.find_all('td'):
            cell_td = str(cell_td.text).strip()
            if ('：' not in cell_td) and (cell_td != '详细信息') and (cell_td != '证书附件') and (cell_td != '未上传附件')\
                    and (cell_td != ''):
                # print(cell_td)
                information_data_list.append(cell_td)

    Information_Data_List.append(information_data_list)

# 按照不同数组长度定义方法，存入不同的数组
First_Data_List_2 = []
def get_first_data_2(url_list):
    html = get_html(url_list[0])
    soup = BeautifulSoup(html, 'html.parser')
    table_table1 = soup.find_all('table', class_='table1 font_s02')
    # print(table_table1)
    for cell_tr in table_table1[0].find_all('tr'):
        first_data_list_2 = []
        first_data_list_2.append(url_list[1])
        first_data_list_2.append(url_list[2])
        for cell_td in cell_tr.find_all('td'):
            cell_td = cell_td.text
            if (cell_td != '序号') and (cell_td != '代码及名称'):
                first_data_list_2.append(cell_td)
        # print(first_data_list_2)
        if len(first_data_list_2) != 2:
            First_Data_List_2.append(first_data_list_2)

# 按照不同数组长度定义方法，存入不同的数组
First_Data_List_4 = []
def get_first_data_4(url_list):
    html = get_html(url_list[0])
    soup = BeautifulSoup(html, 'html.parser')
    table_table1 = soup.find_all('table', class_='table1 font_s02')
    # print(table_table1)
    for cell_tr in table_table1[0].find_all('tr'):
        first_data_list_4 = []
        first_data_list_4.append(url_list[1])
        first_data_list_4.append(url_list[2])
        for cell_td in cell_tr.find_all('td'):
            cell_td = cell_td.text
            if (cell_td != '序号') and (cell_td != '类别代码') and (cell_td != '名称') and (cell_td != '全部/部分业务范围'):
                first_data_list_4.append(cell_td)

        if len(first_data_list_4) != 2:
            First_Data_List_4.append(first_data_list_4)


# 按照不同数组长度定义方法，存入不同的数组
First_Data_List_5 = []
def get_first_data_5(url_list):
    html = get_html(url_list[0])
    soup = BeautifulSoup(html, 'html.parser')
    table_table1 = soup.find_all('table', class_='table1 font_s02')
    # print(table_table1)
    for cell_tr in table_table1[0].find_all('tr'):
        first_data_list_5 = []
        first_data_list_5.append(url_list[1])
        first_data_list_5.append(url_list[2])
        for cell_td in cell_tr.find_all('td'):
            cell_td = cell_td.text
            if (cell_td != '序号') and (cell_td != '产品名称') and (cell_td != '产品标准') and (cell_td != '标准代码') \
                    and (cell_td != '认证模式'):
                first_data_list_5.append(cell_td)

        if len(first_data_list_5) != 2:
            First_Data_List_5.append(first_data_list_5)


# 按照不同数组长度定义方法，存入不同的数组
First_Data_List_2_4 = []
def get_first_data_2_4(url_list):
    html = get_html(url_list[0])
    soup = BeautifulSoup(html, 'html.parser')
    table_table1 = soup.find_all('table', class_='table1 font_s02')
    # print(table_table1)
    for cell_tr in table_table1[0].find_all('tr'):
        first_data_list_2_4 = []
        first_data_list_2_4.append(url_list[1])
        first_data_list_2_4.append(url_list[2])
        i = 0
        for cell_td in cell_tr.find_all('td'):
            i += 1
            cell_td = cell_td.text
            if (cell_td != '序号') and (cell_td != '代码及名称') and (cell_td != ''):
                first_data_list_2_4.append(cell_td)

            if len(first_data_list_2_4) != 2 and (i % 2 == 0):
                First_Data_List_2_4.append(first_data_list_2_4)
                first_data_list_2_4 = []
                first_data_list_2_4.append(url_list[1])
                first_data_list_2_4.append(url_list[2])

# 定义一个返回部分信息的方法，返回含有部分信息的数组second_data_list
Second_Data_List = []
def get_second_data(url_list):
    html = get_html(url_list[0])
    soup = BeautifulSoup(html, 'html.parser')
    table_table1 = soup.find_all('table', class_='table1 font_s02')
    for cell_tr in table_table1[0].find_all('tr'):
        second_data_list = []
        second_data_list.append(url_list[1])
        second_data_list.append(url_list[2])
        for cell_td in cell_tr.find_all('td'):
            cell_td = str(cell_td.text).strip()
            if (cell_td != '序号') and (cell_td != '类别代码') and (cell_td != '业务范围描述'):
                second_data_list.append(cell_td)

        if len(second_data_list) != 2:
            Second_Data_List.append(second_data_list)


# 定义一个存储分支结构信息的方法，存储含有分支信息的数组Third_Data_List
Third_Data_List = []
def get_third_data(url_list):
    html = get_html(url_list[0])
    soup = BeautifulSoup(html, 'html.parser')
    table_table1 = soup.find_all('table', class_='viewTable')
    table_tbody = table_table1[0].find_all('tbody')
    for cell_tr in table_tbody[0].find_all('tr'):
        third_data_list = []
        third_data_list.append(url_list[1])
        third_data_list.append(url_list[2])
        for cell_td in cell_tr.find_all('td'):
            cell_td = str(cell_td.text).strip()
            # print(cell_td)
            third_data_list.append(cell_td)

        if len(third_data_list) != 2:
            Third_Data_List.append(third_data_list)

if __name__ == '__main__':
    url = 'http://cb.cnas.org.cn/cnas_published/servlet/NetShowServlet?' \
          'field=all&address=&certNo=&actionType=showOrgList&orgName=&clazz=&d-49751-p='
    print(url + '1')
    # for num in range(57):
    #     print(num + 1)
    #     get_url_data(url + str(num+1))
    # df_Sheet1 = pd.DataFrame(Name_Data_List, columns=['单位名称', '领域', '有效期', '批准日期', '首次发证日', '负责人',
    #                                                   '联系人', '状态', '是否签署MLA许可协议'])
    # df_Sheet1.to_csv('E:/Pycharm_xjf/JiGou_PaChong/ZhiLiang_pc/简略信息.csv')
    # print('------df_Sheet1 Successful!!!!!!------')
    #
    # page_num = 1
    # for url_str in Url_List:
    #     print(page_num)
    #     get_next_data(url_str)
    #     page_num += 1
    # df_Sheet2 = pd.DataFrame(Information_Data_List, columns=['机构名称', '证书号', '批准时间', '有效期', '负责人', '电话',
    #                                                          '联系人', '地址', '邮编', '传真', '机构网址', '机构Email',
    #                                                          '领域', '首次发证时间', '认可要求'])
    # df_Sheet2.to_csv('E:/Pycharm_xjf/JiGou_PaChong/ZhiLiang_pc/详细信息.csv')
    # print('------df_Sheet2 Successful!!!!!!------')
    #
    # page_num = 1
    # for url_list in First_Url_List_2:
    #     print(page_num)
    #     get_first_data_2(url_list)
    #     page_num += 1
    # df_Sheet3 = pd.DataFrame(First_Data_List_2, columns=['机构名称', '周期', '序号', '代码及名称'])
    # df_Sheet3.to_csv('E:/Pycharm_xjf/JiGou_PaChong/ZhiLiang_pc/业务范围_2.csv')
    # print('------df_Sheet3 Successful!!!!!!------')
    #
    # page_num = 1
    # for url_list in First_Url_List_4:
    #     print(page_num)
    #     get_first_data_4(url_list)
    #     page_num += 1
    # df_Sheet4 = pd.DataFrame(First_Data_List_4, columns=['机构名称', '周期', '序号', '类别代码', '名称', '全部/部分业务范围'])
    # df_Sheet4.to_csv('E:/Pycharm_xjf/JiGou_PaChong/ZhiLiang_pc/业务范围_4.csv')
    # print('------df_Sheet4 Successful!!!!!!------')
    #
    # page_num = 1
    # for url_list in First_Url_List_5:
    #     print(page_num)
    #     get_first_data_5(url_list)
    #     page_num += 1
    # df_Sheet5 = pd.DataFrame(First_Data_List_5, columns=['机构名称', '周期', '序号', '产品名称', '产品标准', '标准代码', '认证模式'])
    # df_Sheet5.to_csv('E:/Pycharm_xjf/JiGou_PaChong/ZhiLiang_pc/业务范围_5.csv')
    # print('------df_Sheet5 Successful!!!!!!------')
    #
    # page_num = 1
    # for url_list in First_Url_List_2_4:
    #     print(page_num)
    #     get_first_data_2_4(url_list)
    #     page_num += 1
    # df_Sheet6 = pd.DataFrame(First_Data_List_2_4, columns=['机构名称', '周期', '序号', '代码及名称'])
    # df_Sheet6.to_csv('E:/Pycharm_xjf/JiGou_PaChong/ZhiLiang_pc/业务范围_2_4.csv')
    # print('------df_Sheet6 Successful!!!!!!------')
    #
    # page_num = 1
    # for url_list in Second_Url_List:
    #     print(page_num)
    #     get_second_data(url_list)
    #     page_num += 1
    # df_Sheet7 = pd.DataFrame(Second_Data_List, columns=['机构名称', '周期', '序号', '类别代码', '业务描述范围'])
    # df_Sheet7.to_csv('E:/Pycharm_xjf/JiGou_PaChong/ZhiLiang_pc/认可的部分业务范围.csv')
    # print('------df_Sheet7 Successful!!!!!!------')
    #
    # page_num = 1
    # for url_list in Third_Url_List:
    #     print(page_num)
    #     get_third_data(url_list)
    #     page_num += 1
    # df_Sheet8 = pd.DataFrame(Third_Data_List, columns=['机构名称', '周期', '分场所名称', '分场所地址', '负责人', '联系人', '联系电话'])
    # df_Sheet8.to_csv('E:/Pycharm_xjf/JiGou_PaChong/ZhiLiang_pc/分支机构.csv')
    # print('------df_Sheet8 Successful!!!!!!------')

    # file_path = 'E:/Pycharm_xjf/JiGou_PaChong/ZhiLiang_pc/ZhiLiang_Data.xlsx'
    # write = pd.ExcelWriter(file_path)
    # df_Sheet1.to_excel(excel_writer=write, sheet_name='简略信息', index=None)
    # df_Sheet2.to_excel(excel_writer=write, sheet_name='详细信息', index=None)
    # df_Sheet3.to_excel(excel_writer=write, sheet_name='业务范围_2', index=None)
    # df_Sheet4.to_excel(excel_writer=write, sheet_name='业务范围_4', index=None)
    # df_Sheet5.to_excel(excel_writer=write, sheet_name='业务范围_5', index=None)
    # df_Sheet6.to_excel(excel_writer=write, sheet_name='业务范围_2_4', index=None)
    # df_Sheet7.to_excel(excel_writer=write, sheet_name='认可的部分业务范围', index=None)
    # df_Sheet8.to_excel(excel_writer=write, sheet_name='分支机构', index=None)
    # write.save()
    # print('--------Excel Save Successful!!!!!!--------')
    # write.close()
    # print('--------Excel Close Successful!!!!!!-------')