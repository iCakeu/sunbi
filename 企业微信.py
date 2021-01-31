#! /usr/bin/python
# coding:utf-8

'''
企业微信机器人及API
1 企业微信机器人
2 企业微信API
'''

from 配置.py import configRobot,configQywx
from requests import post 
from hashlib import md5
from base64 import b64encode
from requests_toolbelt import MultipartEncoder
import os
import json 

class 企业微信机器人:
    def _init__(self,口令=configRobot['默认'])
        '''
        请设置一个默认的webhook
        '''
        self.__url = 口令

    def 发文本(self,文本):
        '''
        文本参数可以是
        1 循环型(列表,元组,集合)
        2 文本
        3 其他会被文本化
        '''
        self.__text = 文本
        if isinstance(文本,(list,tuple,set)):
            for text in 文本:
                self._sendText(text)
        elif isinstance(文本,str):
            self._sendText(文本)
        else:
            text = str(文本)
            self._sendText(text)

    def _sendText(self,text):
        self.__data = json.dumps({
            'msgtype':'text',
            'text'   :{
                'contect':text
            }
        })
        post(self.__url,self.__data)  

    def 发Markdown(self,Markdown):
        self.__Markdown = Markdown
        if isinstance(Markdown,(list,tuple,set)):
            for md in Markdown:
                self._sendMarkdown(md)
        else:
            self._sendMarkdown(Markdown)

    def _sendMarkdown(self,mkd):
        self.__data = json.dumps({
            'msgtype':'markdown',
            'markdown':{
                'content':mkd
                }
        })
        post(self.__url,self.__data) 

    def 发图片(self,图片):
        '''
        参数为图片的完整路径 或者完整路径的集合类元素
        '''
        self.__image = 图片
        if isinstance(图片,(list,tuple,set)):
            for img in 图片:
                self._sendImage(img)
        elif isinstance(图片,str):
            self._sendImage(图片)
        else:
            print('请传入完整的图片路径')

    def _sendImage(self,imgPath):
        # md5编码
        # base64编码
        with open(imgPath,'rb') as f:
            m = md5()
            m.update(f.read())
            md = m.hexdigest()
            base = str(b64encode(f.read()),'utf-8')
        
        self.__data = json.dumps({
            'msgtype':'image',
            'image':{
                'base64':base,
                'md5':md
            }
        })
        post(self._url,self.__data)

    def 发图文(self,图文):
        '''
        参数为 列表/元组/集合 内嵌字典的形式 或者直接为字典
        参照微信官方说明看看key值有啥
        自带批量发送 故不用建子类
        '''
        self.__图文 = 图文
        self.__data = json.dumps({
            "msgtype": "news",
            "news": {
            "articles" : 图文
            }
        })
        post(self._url,self.__data)

    def 发文件(self,文件):
        '''
        文件列表或者单个文件完整路径
        '''
        self.__file = 文件
        if isinstance(文件,(list,tuple,set)):
            for file in 文件:
                self._sendFile(file)
        elif isinstance(文件,str):
            self._sendFile(文件)
        else:
            print('请传入完整的文件路径')

    def _sendFile(self,file,key=self.__url[-36:]):
        with open(file, 'rb') as f:
            files = {'uploadFile':(os.path.basename(file),f)}
        data = MultipartEncoder(files)
        headers = {
            'Content-Type':data.content_type,
            'Content-Disposition':'form-data;name="media";filename="test.xlsx";filelength=6'
        }
        __urlUpload = 'https://qyapi.weixin.qq.com/cgi-bin/webhook/upload_media?key=%s&type=file' % key
        response = post(__urlUpload,headers=headers,data=data)
        mid = json.loads(response.text)['media_id']
        
        self.__data = json.dumps({
            'msgtype':'file',
            'file':{
                'media_id':mid
            }
        })
        post(self._url,self.__data)
            
        
class 企业微信API:
    def __init__(self,企业ID=configQywx['企业ID'],客户secret=configQywx['客户secret'],通讯录secret=configQywx['通讯录secret']):
        pass

    def 获取架构(self):
        pass

    def 获取架构成员(self):
        pass

    def 获取成员客户(self):
        pass

    def 获取客户(self):
        pass

    
