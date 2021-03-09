'''
业务场景 
主要使用模块应该在这里
'''

## 1. 网页爬虫
def 网页爬虫():
    '''
    此部分需要自行修改 仅demo用
    '''
    from web import 网页模拟

    下载文件夹 = r'c:\selenium_test'
    网址 = 'https://cn.bing.com'

    网页 = 网页模拟(下载文件夹=下载文件夹)
    网页.访问地址(网址=网址)

    ## 多元素
    网页.点击元素(['//*[@id="fm1"]/div[2]/div[1]','//*[@id="fm1"]/div[2]/div[2]'])
    网页.输入内容(['//*[@id="username"]','//*[@id="password"]'],['danzhao','123'])
    网页.按回车(['//*[@id="fm1"]/div[1]/div[8]/button','//*[@id="fm1"]/div[1]/div[8]/button/button'])

    ## 单元素
    网页.点击元素('//*[@id="fm1"]/div[2]/div[2]')
    网页.输入内容('//*[@id="username"]','danzhao')
    网页.按回车('//*[@id="fm1"]/div[1]/div[8]/button')

    网页.关闭浏览器()

# 网页爬虫()


## 2. api数据传输
def API取数入库(api配置,数据库配置):
    '''
    从api获取数据 然后入库
    '''
    from web import API
    from database import 数据库

    api = API(地址=api配置['地址'],抬头=api配置['抬头'],数据包=api配置['数据包'])
    result = api[api配置['结果key']]
    # result如果是字典形式 需要转换成列表或元组嵌套形式
    listTuple = []
    for dict in result:
        lst = [dict[key] for key in dict]
        listTuple.append(lst)
    
    if len(listTuple) > 0:
        数据库 = 数据库(配置=数据库配置)
        if '开始日期' in 数据库配置:
            sql = 'delete from {table} where {column} between "{dateBegin}" and "{dateEnd}"'
            数据库.跑(脚本或存储过程名=sql.format(
                                    table=数据库配置['表名'],
                                    columns=数据库配置['日期字段'],
                                    dateBegin=数据库配置['开始日期'],
                                    dateEnd=数据库配置['结束日期'],
                                    ),脚本类型='',结束后关闭='否')
            数据库.插(表名=数据库配置['表名'],数据=listTuple)
        else:
            sql = 'delete from {table}'
            数据库.跑(脚本或存储过程名=sql.format(table=数据库配置['表名']),脚本类型='',结束后关闭='否')
            数据库.插(表名=数据库配置['表名'],数据=listTuple)

# api配置 = {
#     '地址' : 'http://10.250.90.117:666',
#     '抬头' : {
#         'cookie':'123asd',
#         'length':128
#     },
#     '数据包' : {
#         'startDate':'2021-03-03',
#         'endDate'  :'2021-03-04'
#     },
#     '结果key':'result'
# }
# 数据库配置 = {
#     '地址'  :'10.240.90.117',
#     '账号'  :'shangde',
#     '密码'  :'shangde',
#     '数据库':'shangde',
#     '端口'  :3306,
#     '表名'  :'dim_date',
#     '日期字段':'everyday',
#     '开始日期':'2021-01-01',
#     '结束日期':'2021-01-02'
# }
# API取数入库(api配置=api配置,数据库配置=数据库配置)


## 3. 数据库同步
def 数据库同步(原数据库,目标数据库,同步配置):
    from database import 数据库
    数据源 = 数据库(配置=原数据库)
    目标库 = 数据库(配置=目标数据库)
    try:
        for d in 同步配置:
            数据   = 数据源.跑(脚本或存储过程名=d['脚本'],结束后关闭='否')
            目标库.插(表名=d['表名'],数据=数据,结束后关闭='否')
    except Exception as e:
        print(e)
    finally:
        数据源.关闭()
        目标库.关闭()
    
# 原数据库 = {
#     '地址'  :'10.240.90.117',
#     '账号'  :'shangde',
#     '密码'  :'shangde',
#     '数据库':'shangde',
#     '端口'  :3306
# }
# 目标数据库 = {
#     '地址'  :'10.240.90.179',
#     '账号'  :'shangde',
#     '密码'  :'shangde',
#     '数据库':'shangde',
#     '端口'  :3306
# }
# 同步配置 = [
#     {
#         '脚本':'select * from dim_date',
#         '表名':'t_date'
#     },
#     {
#         '脚本':'select * from dim_date where 日期 >= "2021-12-01"',
#         '表名':'t_date'
#     }
# ]
# 数据库同步(原数据库=原数据库,目标数据库=目标数据库,同步配置=同步配置)


## 4. ftp文件传输入库
def ftp文件传输(ftp配置,数据库配置):
    '''
    无需登录的ftp
    适合ftp内只有单一的文件结构
    '''
    from file import 文件
    from database import 数据库
    文件 = 文件(文件夹=ftp配置['文件夹'])
    文件.ftp复制文件(ftpip=ftp配置['ftpip'],目录=ftp配置['目录'])
    文件列表 = 文件.获取文件夹文件()
    if len(文件列表) > 0:
        数据库 = 数据库(配置=数据库配置)
        isDel = False
        try:
            for file in 文件列表:
                df = pandas.read_excel(file)
                df.fillna(value='',inplace=True)
                listTuple = df.values.tolist()
                # 删除日期只执行一次
                if '开始日期' in 数据库配置:
                    if not isDel:
                        sql = 'delete from {table} where {column} between "{dateBegin}" and "{dateEnd}"'
                        数据库.跑(脚本或存储过程名=sql.format(
                                                table=数据库配置['表名'],
                                                columns=数据库配置['日期字段'],
                                                dateBegin=数据库配置['开始日期'],
                                                dateEnd=数据库配置['结束日期'],
                                                ),脚本类型='',结束后关闭='否')
                        
                    数据库.插(表名=数据库配置['表名'],数据=listTuple)
                    isDel = True
                else:
                    sql = 'delete from {table}'
                    数据库.跑(脚本或存储过程名=sql.format(table=数据库配置['表名']),脚本类型='',结束后关闭='否')
                    数据库.插(表名=数据库配置['表名'],数据=listTuple)
        except Exception as e:
            print(e)
        finally:
            数据库.关闭()

# ftp配置 = {
#     '文件夹':r'C:\file_test',
#     'ftpip':'10.247.63.82',
#     '目录':'headquarter'
# }
# 数据库配置 = {
#     '地址'  :'10.240.90.117',
#     '账号'  :'shangde',
#     '密码'  :'shangde',
#     '数据库':'shangde',
#     '端口'  :3306,
#     '表名'  :'dim_date',
#     '日期字段':'everyday',
#     '开始日期':'2021-01-01',
#     '结束日期':'2021-01-02'
# }
# 文件列表 = ftp文件传输(ftp配置=ftp配置,数据库配置=数据库配置)


## 5. 邮件爬取
def 下载邮件附件(邮件配置):
    '''
    在收件箱找到指定邮件 下载附件
    '''
    from mail import 邮件
    from file import 文件
    文件 = 邮件配置['文件夹']
    邮件 = 邮件(邮件配置=邮件配置)
    邮件.按主题下载附件(
        关键字=邮件配置['关键字'],
        保存文件夹=邮件配置['文件夹'],
        第几封=邮件配置['第几封'],
        找到多少封后停止=邮件配置['第几封停止']
        )
    return 文件.获取文件夹文件()

# 邮件配置 = {
#     'username':'shangde',
#     'password':'shangde',
#     '关键字':'订单',
#     '文件夹':r'c:\mail_test',
#     '第几封':1,
#     '第几封停止':20
# }
# 文件列表 = 下载邮件附件(邮件配置=邮件配置)


## 6. 跑存储过程,刷新Excel,分发
### (存储过程,脚本) x (刷新,运行宏) x (邮件分发,企业微信分发) = 4
### Excel (取图 x 取文字 x 取文件*)
def 脚本刷新分发(表格配置,分发配置,数据库配置=''):
    from database import 数据库
    from excel import 表格
    from mail import 邮件 
    from workWechat import 企业微信机器人
    import os

    # 存储过程/脚本
    if 数据库配置 != '':
        数据库 = 数据库(配置=数据库配置)
        数据库.跑(脚本或存储过程名=数据库配置['脚本'],脚本类型=数据库配置['脚本类型'],存储过程参数=数据库配置['存储过程参数'],结束后关闭='是')

    # Excel
    attach = []
    text   = []
    pic    = []
    for 表 in 表格配置:
        表格 = 表格(工作簿路径=表['工作簿路径'])
        ## 宏or刷新 二选一
        if '宏' in 表:
            表格.运行宏另存为关闭(宏名=表['宏'],另存为路径=表['另存为路径'],休眠时间=表['休眠时间'])
        else:
            表格.刷新另存为关闭(另存为路径=表['另存为路径'],休眠时间=表['休眠时间'])
        attach.append(表['另存为路径'])

        ## 文本
        if '文本' in 表:
            表格 = 表格(工作簿路径=表['另存为路径'])
            text.append(表格.获取单元格文本(行列=表['文本']))
            表格.关闭()
        
        ## 图片
        if '图片' in 表:
            表格 = 表格(工作簿路径=表['另存为路径'])
            pic.append(表格.生成图片(
                    对象=os.path.splitext(os.path.basename(表['图片']))[0],
                    文件夹=os.path.split(表['图片'])[0],
                    格式=os.path.splitext(os.path.basename(表['图片']))[-1]
                    ))
            attach.extend(表['图片'])
            表格.关闭()

    # 分发
    if '邮件' in 分发配置:
        邮件配置 = 分发配置['邮件']
        邮件 = 邮件(邮件配置=邮件配置)
        邮件.发(
            主题=邮件配置['主题'],
            收件人=邮件配置['收件人'],
            正文=邮件配置['正文'],
            附件=attach,
            抄送=邮件配置['抄送']
            )

    if '企业微信' in 分发配置:
        企业微信配置 = 分发配置['企业微信']
        for 人 in 企业微信配置:
            机器人 = 企业微信机器人(口令=人['口令'])
            机器人.发文件(文件=attach)
            if '文本' in 人:
                机器人.发文本(文本=人['文本'])
            if '图片' in 人:
                机器人.发图片(图片=人['图片'])

# 数据库配置 = {
#     '地址'  :'10.240.90.117',
#     '账号'  :'shangde',
#     '密码'  :'shangde',
#     '数据库':'shangde',
#     '端口'  :3306,
#     '脚本'  :'dim_oid_order',   # 可以是存储过程名 也可以直接是脚本 根据脚本类型决定
#     '脚本类型':'存储过程',
#     '存储过程参数':()
# }
# 表格配置 = [
#     {
#         '工作簿路径':r'C:\excel_test\test1.xlsx',
#         '另存为路径':r'C:\excel_test\test20210101.xlsx',
#         '休眠时间':10,
#         '文本':((1,2),),
#         '图片':[r'c:\excel_test\a.jpg'],
#         '宏':'dan'
#     },
#     {
#         '工作簿路径':r'C:\excel_test\test2.xlsx',
#         '另存为路径':r'C:\excel_test\test20210102.xlsx',
#         '休眠时间':10
#     }
# ]
# 分发配置 = {
#     '邮件' : {
#         '账号':'shangde',
#         '密码':'shangde',
#         '主题':'测试主题',
#         '收件人':[('但老师1','danzhao@sunlands.com'),('但老师2','danzhao@sunlands.com')],
#         '正文':'''
#             <ul>
#                 <li>列表</li>
#                 <li>列表</li>
#             </ul>
#             ''',
#         '抄送':[('但老师1','danzhao@sunlands.com'),('但老师2','danzhao@sunlands.com')]
#     },
#     '企业微信' : [
#         {
#             '口令':'123',
#             '文本':['abc','bcd'],
#             '图片':[r'c:\qytest\a.jpg',r'c:\qytest\b.jpg']
#         },
#         {
#             '口令':'123',
#             '文本':['abc','bcd'],
#             '图片':[r'c:\qytest\a.jpg',r'c:\qytest\b.jpg']
#         }
#     ]
# }
# 脚本刷新分发(数据库配置=数据库配置,表格配置=表格配置,分发配置=分发配置)


## 7 脚本分发
### 脚本 x 导出(写文件) x (邮件,企业微信) = 2
def 脚本导出分发(数据库配置,分发配置):
    from database import 数据库
    from excel import 表格
    from mail import 邮件 
    from workWechat import 企业微信机器人

    # 存储过程/脚本
    数据库 = 数据库(配置=数据库配置)
    数据 = 数据库.跑(脚本或存储过程名=数据库配置['脚本'],脚本类型='查询',结束后关闭='是')

    # Excel
    excel = 表格().导出(文件路径=数据库配置['文件'],嵌套对象=数据,标题=数据库配置['标题'])

    # 分发
    if '邮件' in 分发配置:
        邮件配置 = 分发配置['邮件']
        邮件 = 邮件(邮件配置=邮件配置)
        邮件.发(
            主题=邮件配置['主题'],
            收件人=邮件配置['收件人'],
            正文=邮件配置['正文'],
            附件=excel,
            抄送=邮件配置['抄送']
            )

    if '企业微信' in 分发配置:
        企业微信配置 = 分发配置['企业微信']
        for 人 in 企业微信配置:
            机器人 = 企业微信机器人(口令=人['口令'])
            机器人.发文件(文件=excel)

# 数据库配置 = {
#     '地址'  :'10.240.90.117',
#     '账号'  :'shangde',
#     '密码'  :'shangde',
#     '数据库':'shangde',
#     '端口'  :3306,
#     '脚本'  :'select * from dim_date',
#     '文件'  :r'c:\file_test\test.xlsx',
#     '标题'  :['日期','年','月','周','日']
# }
# 分发配置 = {
#     '邮件' : {
#         '账号':'shangde',
#         '密码':'shangde',
#         '主题':'测试主题',
#         '收件人':[('但老师1','danzhao@sunlands.com'),('但老师2','danzhao@sunlands.com')],
#         '正文':'''
#             <ul>
#                 <li>列表</li>
#                 <li>列表</li>
#             </ul>
#             ''',
#         '抄送':[('但老师1','danzhao@sunlands.com'),('但老师2','danzhao@sunlands.com')]
#     },
#     '企业微信' : [
#         {
#             '口令':'123'
#         },
#         {
#             '口令':'123'
#         }
#     ]
# }
# 脚本导出分发(数据库配置=数据库配置,分发配置=分发配置)


## 8 脚本分发
### 脚本 x 导出(纯文本) x 邮件 
### html table
def 脚本导出邮件(分发配置,数据库配置):
    from database import 数据库
    from mail import 邮件 

    # 存储过程/脚本
    数据库 = 数据库(配置=数据库配置)
    数据 = 数据库.跑(脚本或存储过程名=数据库配置['脚本'],脚本类型='查询',结束后关闭='是')

    # html
    body = '''
        <table style='text-align:center;border-collapse:collapse;font-size:14px;width:100%;'>
            <thead>
                {head}
            </thead>
            <tbody>
                {cont}
            </tbody>
        </table>
    '''
    head = "<tr style='text-align:center;height:50px;'><th>{title}</th></tr>".format(title='</th><th>'.join(数据库配置['标题']))
    cont = []
    for x in 数据:
        element = '<tr style='text-align:center;height:50px;'><td>{ele}</td></tr>'.format(ele='</td><td>'.join(x))
        cont.append(element)
    cont = '\n'.join(cont)

    # 分发
    if '邮件' in 分发配置:
        邮件配置 = 分发配置['邮件']
        邮件 = 邮件(邮件配置=邮件配置)
        邮件.发(
            主题=邮件配置['主题'],
            收件人=邮件配置['收件人'],
            正文=body.format(head=head,cont=cont),
            抄送=邮件配置['抄送']
            )

# 数据库配置 = {
#     '地址'  :'10.240.90.117',
#     '账号'  :'shangde',
#     '密码'  :'shangde',
#     '数据库':'shangde',
#     '端口'  :3306,
#     '脚本'  :'select * from dim_date',
#     '文件'  :r'c:\file_test\test.xlsx',
#     '标题'  :['日期','年','月','周','日']
# }
# 邮件配置 = {
#     '账号':'shangde',
#     '密码':'shangde',
#     '主题':'测试主题',
#     '收件人':[('但老师1','danzhao@sunlands.com'),('但老师2','danzhao@sunlands.com')],
#     '抄送':[('但老师1','danzhao@sunlands.com'),('但老师2','danzhao@sunlands.com')]
# }
# 脚本导出邮件(数据库配置=数据库配置,分发配置=分发配置)
