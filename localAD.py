#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@Author: randolph
@Date: 2020-05-27 14:33:03
@LastEditors: randolph
@LastEditTime: 2020-06-11 15:22:39
@version: 2.0
@Contact: cyg0504@outlook.com
@Descripttion: 用python3+ldap3管理windows server2019的AD域;
'''
import json
import logging.config
import os
import random
import re
import string
from time import sleep

import pandas as pd
import winrm
import yaml
from ldap3 import (ALL, ALL_ATTRIBUTES, MODIFY_REPLACE, NTLM, SASL, SIMPLE,
                   SUBTREE, SYNC, Connection, Server, LEVEL)
from tqdm import tqdm

# 日志配置
LOG_CONF = 'logging.yaml'
# AD域设置
LDAP_IP = '192.168.255.223'                                 # LDAP本地服务器IP
USER = 'CN=Administrator,CN=Users,DC=randolph,DC=com'       # LDAP本地服务器IP
PASSWORD = "QQqq#123"                                       # LDAP本地服务器管理员密码

DISABLED_BASE_DN = 'OU=resigned,DC=randolph,DC=com'        # 离职账户所在OU
ENABLED_BASE_DN = "OU=上海总部,DC=randolph,DC=com"         # 正式员工账户所在OU
USER_SEARCH_FILTER = '(objectclass=user)'                  # 只获取用户对象 过滤条件
OU_SEARCH_FILTER = '(objectclass=organizationalUnit)'      # 只获取OU对象 过滤条件
DISABLED_USER_FLAG = [514, 546, 66050, 66080, 66082]       # 禁用账户UserAccountControl对应十进制值列表
ENABLED_USER_FLAG = [512, 544, 66048, 262656]              # 启用账户UserAccountControl对应十进制值列表
# excel表格
RAN_EXCEL = "ran_list.xlsx"                                 # 原始造的数据
TEST_RAN_EXCEL = "test_ran_list.xlsx"                       # 测试用表格
NEW_RAN_EXCEL = "new_ran_list.xlsx"                         # 新增员工表格
PWD_PATH = 'pwd.txt'
# WINRM信息 无需设置
WINRM_USER = 'Administrator'
WINRM_PWD = PASSWORD


class AD(object):
    '''AD域的操作
    '''

    def __init__(self):
        '''初始化加载日志配置
        AD域连接
        AD基础信息加载
        '''
        # 初始化加载日志配置
        self.__setup_logging(path=LOG_CONF)
        SERVER = Server(host=LDAP_IP,
                        port=636,               # 636安全端口
                        use_ssl=True,
                        get_info=ALL,
                        connect_timeout=3)      # 连接超时为3秒
        try:
            self.conn = Connection(
                server=SERVER,
                user=USER,
                password=PASSWORD,
                auto_bind=True,
                read_only=False,                # 禁止修改数据True
                receive_timeout=10)             # 10秒内没返回消息则触发超时异常
            logging.info("distinguishedName:%s res: %s" % (USER, self.conn.bind()))
        except BaseException as e:
            logging.error("AD域连接失败，请检查IP/账户/密码")
        finally:
            self.conn.closed

    def __setup_logging(self, path=LOG_CONF, default_level=logging.INFO, env_key="LOG_CFG"):
        value = os.getenv(env_key, None)
        if value:
            path = value
        if os.path.exists(path):
            with open(path, "r") as f:
                config = yaml.safe_load(f)
                logging.config.dictConfig(config)
        else:
            logging.basicConfig(level=default_level)

    def get_users(self, attr=ALL_ATTRIBUTES):
        '''
        @param {type}
        @return: total_entries 此次查询到的记录数目
        @msg: 获取所有用户
        '''
        entry_list = self.conn.extend.standard.paged_search(
            search_base=ENABLED_BASE_DN,
            search_filter=USER_SEARCH_FILTER,
            search_scope=SUBTREE,
            attributes=attr,
            paged_size=5,
            generator=False)                    # 关闭生成器，结果为列表
        total_entries = 0
        for entry in entry_list:
            total_entries += 1
        logging.info("共查询到记录条目: " + str(total_entries))
        return entry_list

    def get_ous(self, attr=None):
        '''
        @param {type}
        @return: res所有OU
        @msg: 获取所有OU
        '''
        self.conn.search(search_base=ENABLED_BASE_DN,
                         search_filter=OU_SEARCH_FILTER,
                         attributes=attr)
        result = self.conn.response_to_json()
        res_list = json.loads(result)['entries']
        return res_list[::-1]
    
    def get_level_users(self, SEARCH_BASE, attr=ALL_ATTRIBUTES):
        '''
        @param {type}
        @return: total_entries 此次查询到的记录数目
        @msg: 获取某OU下用户
        '''
        entry_list = self.conn.extend.standard.paged_search(
            search_base=SEARCH_BASE,
            search_filter=USER_SEARCH_FILTER,
            search_scope=LEVEL,
            attributes=attr,
            paged_size=5,
            generator=False)                    # 关闭生成器，结果为列表
        total_entries = 0
        for entry in entry_list:
            total_entries += 1
        logging.info("共查询到记录条目: " + str(total_entries))
        return entry_list
    
    def handle_excel(self, path):
        '''
        @param path{string} excel文件绝对路径
        @return: result: { 'page_flag': True, 'person_list': [[], [], ...] }
        @msg: 表格文件预处理
        1.增加行列数判————行数决定AD的查询是否分页，列数用以判断必须列数据完整性与补充列；
        2.判断必须列【工号|姓名|部门】是否存在且是否有空值
        3.人员列表的使用sort函数排序key用lambda函数，排序条件(i[2].count('.'), i[2], i[0])为(部门层级、部门名称、工号)
        '''
        try:
            # 1.开始源文件格式扫描
            df = pd.read_excel(path, encoding='utf-8', error_bad_lines=False)           # 读取源文件
            a, b = df.shape                                                             # 表格行列数
            cols = df.columns.tolist()                  # 表格列名列表
            is_ex_null = df.isnull().any().tolist()     # 列是否存在空值
            dic = dict(zip(cols, is_ex_null))           # 存在空值的列
            if int("工号" in cols) + int("姓名" in cols) + int("部门" in cols) < 3:     # 判断必须列是否都存在
                logging.error("表格缺少必要列【工号|姓名|部门】请选择正确的源文件;或者将相应列列名修改为【工号|姓名|部门】")
                exit()
            elif int(dic["工号"]) + int(dic["姓名"]) + int(dic["部门"]) > 0:            # 判断必须列是否有空值
                logging.error("必要列存在空值记录，请检查补全后重试：" + '\n' + str(df[df.isnull().values == True]))
            else:
                df = pd.read_excel(path, encoding='utf-8', error_bad_lines=False, usecols=[i for i in range(0, b)])
                use_cols = ["工号", "姓名", "部门"]     # 使用的必须列
                for c in ["邮件", "电话", "岗位"]:      # 扩展列的列名在这里添加即可
                    if c in cols:
                        use_cols.append(c)
                df = df[use_cols]                       # 调整df使用列顺序
                person_list = df.values.tolist()        # df数据框转list
                person_list.sort(key=lambda i: (i[2].count('.'), i[2], i[0]), reverse=False)        # 多条件排序
                # 2.开始处理列表
                for i, row in enumerate(person_list):
                    job_id, name, depart = row[0:3]
                    # 将部门列替换成DN
                    row[2] = 'CN=' + str(name + str(job_id)) + ',' + 'OU=' + ',OU='.join(row[2].split('.')[::-1]) + ',' + ENABLED_BASE_DN
                    row.append('RAN' + str(job_id).zfill(6))        # 增加登录名列,对应AD域user的 sAMAccountname 属性
                    row.append(name + str(job_id))                  # 增加CN列,对应user的 cn 属性
                # 3.开始处理返回字典
                result_dic = dict()                         # 返回字典
                if a > 1000:
                    result_dic['page_flag'] = True
                else:
                    result_dic['page_flag'] = False
                result_dic['person_list'] = person_list
                return result_dic
        except Exception as e:
            logging.error(e)
            return None

    def generate_pwd(self, count):
        '''
        @param count{int} 所需密码长度
        @return: pwd: 生成的随机密码
        @msg: 生成随机密码，必有数字、大小写、特殊字符且数目伪均等；
        '''
        pwd_list = []
        a, b = count // 4, count % 4
        # 四种类别先均分除数个字符
        pwd_list.extend(random.sample(string.digits, a))
        pwd_list.extend(random.sample(string.ascii_lowercase, a))
        pwd_list.extend(random.sample(string.ascii_uppercase, a))
        pwd_list.extend(random.sample('!@#$%^&*()', a))
        # 从四种类别中再取余数个字符
        pwd_list.extend(random.sample(string.digits + string.ascii_lowercase + string.ascii_uppercase + '!@#$%^&*()', b))
        random.shuffle(pwd_list)
        pwd_str = ''.join(pwd_list)
        return pwd_str

    def write2txt(self, path, content):
        '''
        @param path{string} 写入文件路径;content{string} 每行写入内容
        @return:
        @msg: 每行写入文件
        '''
        try:
            if os.path.exists(path):
                with open(path, mode='a', encoding='utf-8') as file:
                    file.write(content + '\n')
            else:
                with open(path, mode='a', encoding='utf-8') as file:
                    file.write(content + '\n')
            return True
        except Exception as e:
            logging.error(e)
            return False

    def del_ou_right(self, flag):
        '''
        @param cmd_l{list} 待执行的powershell命令列表
        @return: True/False
        @msg: 连接远程windows并批量执行powershell命令
        '''
        # powershell命令 用于启用/关闭OU 防止对象被意外删除 属性
        # 防止对象被意外删除×
        enable_del = ["Import-Module ActiveDirectory",
                      "Get-ADOrganizationalUnit -filter * -Properties ProtectedFromAccidentalDeletion | where {"
                      "$_.ProtectedFromAccidentalDeletion -eq $true} |Set-ADOrganizationalUnit "
                      "-ProtectedFromAccidentalDeletion $false"]
        # 防止对象被意外删除√
        disable_del = ["Import-Module ActiveDirectory",
                       "Get-ADOrganizationalUnit -filter * -Properties ProtectedFromAccidentalDeletion | where {"
                       "$_.ProtectedFromAccidentalDeletion -eq $false} |Set-ADOrganizationalUnit "
                       "-ProtectedFromAccidentalDeletion $true"]
        flag_map = {0: enable_del, 1: disable_del}

        try:
            win = winrm.Session('http://' + LDAP_IP + ':5985/wsman', auth=(WINRM_USER, WINRM_PWD))
            for cmd in flag_map[flag]:
                ret = win.run_ps(cmd)
            if ret.status_code == 0:      # 调用成功 减少日志写入
                # if flag == 0:
                #     logging.info("防止对象被意外删除×")
                # elif flag == 1:
                #     logging.info("防止对象被意外删除√")
                return True
            else:
                return False
        except Exception as e:
            logging.error(e)
            return False

    def create_obj(self, dn=None, type='user', info=None):
        '''
        @param dn{string}, type{string}'user'/'ou'
        @return: res新建结果, self.conn.result修改结果
        @msg:新增对象
        '''
        object_class = {'user': ['user', 'posixGroup', 'top'],
                        'ou': ['organizationalUnit', 'posixGroup', 'top'],
                        }
        if info is not None:
            [job_id, name, dn, email, tel, title, sam, cn] = info
            user_attr = {'sAMAccountname': sam,      # 登录名
                         'userAccountControl': 544,  # 启用账户
                         'title': title,             # 头衔
                         'givenName': name[0:1],     # 姓
                         'sn': name[1:],             # 名
                         'displayname': name,        # 姓名
                         'mail': email,              # 邮箱
                         'telephoneNumber': tel,     # 电话号
                         }
        else:
            user_attr = None
        # 创建之前需要对dn中的OU部分进行判断，如果没有需要创建
        dn_base = dn.split(',', 1)[1]
        check_ou_res = self.check_ou(dn_base)
        if not check_ou_res:
            logging.error('check_ou失败，未知原因！')
            return False
        else:
            self.conn.add(dn=dn, object_class=object_class[type], attributes=user_attr)
            add_result = self.conn.result

            if add_result['result'] == 0:
                logging.info('新增对象【' + dn + '】成功!')
                if type == 'user':          # 若是新增用户对象，则需要一些初始化操作
                    self.conn.modify(dn, {'userAccountControl': [('MODIFY_REPLACE', 512)]})         # 激活用户                                                               # 如果是用户时
                    new_pwd = self.generate_pwd(8)
                    old_pwd = ''
                    self.conn.extend.microsoft.modify_password(dn, new_pwd, old_pwd)                # 初始化密码
                    info = 'DN: ' + dn + ' PWD: ' + new_pwd
                    save_res = self.write2txt(PWD_PATH, info)                                       # 将账户密码写入文件中
                    if save_res:
                        logging.info('保存初始化账号密码成功！')
                    else:
                        logging.error('保存初始化账号密码失败: ' + info)
                    self.conn.modify(dn, {'pwdLastSet': (2, [0])})                                  # 设置第一次登录必须修改密码
            elif add_result['result'] == 68:
                logging.error('entryAlreadyExists 用户已经存在')
            elif add_result['result'] == 32:
                logging.error('noSuchObject 对象不存在ou错误')
            else:
                logging.error('新增对象: ' + dn + ' 失败！其他未知错误')
            return add_result

    def del_obj(self, dn, type):
        '''
        @param dn{string}
        @return: res修改结果
        @msg: 删除对象
        '''
        if type == 'ou':
            self.del_ou_right(flag=0)
            res = self.conn.delete(dn=dn)
            self.del_ou_right(flag=1)
        else:
            res = self.conn.delete(dn=dn)
        if res == True:
            logging.info('删除对象' + dn + '成功！')
            return res
        else:
            return False

    def update_obj(self, old_dn, info=None):
        '''
        @param {type}
        @return:
        @msg: 更新对象
        '''
        if info is not None:
            [job_id, name, dn, email, tel, title, sam, cn] = info
            # 组成更新属性之前需要对dn中的OU部分进行判断，如果没有需要创建
            dn_base = dn.split(',', 1)[1]
            check_ou_res = self.check_ou(dn_base)
            if not check_ou_res:
                logging.error('check_ou失败，未知原因！')
                return False
            else:
                attr = {'distinguishedName': dn,    # dn
                        'sAMAccountname': sam,      # 登录名
                        'title': title,             # 头衔
                        'givenName': name[0:1],     # 姓
                        'sn': name[1:],             # 名
                        'displayname': name,        # 姓名
                        'mail': email,              # 邮箱
                        'telephoneNumber': tel,     # 电话号
                        }
        else:
            attr = None
        changes_dic = {}
        for k, v in attr.items():
            if not self.conn.compare(dn=old_dn, attribute=k, value=v):                  # 待修改属性
                if k == "distinguishedName":        # 若属性有distinguishedName则需要移动user或ou
                    # 若dn修改了需要将密码文件这个人的dn信息更新下
                    self.update_pwd_file_line(old_dn=old_dn, new_dn=dn)
                    self.move_obj(dn=old_dn, new_dn=v)
                changes_dic.update({k: [(MODIFY_REPLACE, [v])]})
        if len(changes_dic) != 0:   # 有修改的属性时
            modify_res = self.conn.modify(dn=dn, changes=changes_dic)
            logging.info('更新对象: ' + dn + ' 更新内容: ' + str(changes_dic))
        return self.conn.result

    def rename_obj(self, dn, newname):
        '''
        @param newname{type}新的名字，User格式："cn=新名字";OU格式："OU=新名字"
        @return: 修改结果
        @msg: 重命名对象
        '''
        res = self.conn.modify_dn(dn, newname)
        if res == True:
            return True
        else:
            return False

    def move_obj(self, dn, new_dn):
        '''
        @param {type}
        @return:
        @msg: 移动对象到新OU
        '''
        relative_dn, superou = new_dn.split(",", 1)
        res = self.conn.modify_dn(dn=dn, relative_dn=relative_dn, new_superior=superou)
        if res == True:
            return True
        else:
            return False

    def compare_attr(self, dn, attr, value):
        '''
        @param {type}
        @return:
        @msg:比较员工指定的某个属性
        '''
        res = self.conn.compare(dn=dn, attribute=attr, value=value)
        return res

    def check_ou(self, ou, ou_list=None):
        '''
        @param {type}
        @return:
        @msg: 递归函数
    如何判断OU是修改了名字而不是新建的：当一个OU里面没有人就判断此OU被修改了名字，删除此OU；
    不管是新建还是修改了名字，都会将人员转移到新的OU下面：需要新建OU则创建OU后再添加/转移人员
    check_ou的作用是为人员的变动准备好OU
        '''
        if ou_list is None:
            ou_list = []
        self.conn.search(ou, OU_SEARCH_FILTER)      # 判断OU存在性

        while self.conn.result['result'] == 0:
            if ou_list:
                for ou in ou_list[::-1]:
                    self.conn.add(ou, 'organizationalUnit')
            return True
        else:
            ou_list.append(ou)
            ou = ",".join(ou.split(",")[1:])
            self.check_ou(ou, ou_list)  # 递归判断
            return True

    def scan_ou(self):
        '''扫描的时候，必须保证此OU为叶子节点，否则报notAllowedOnNonLeaf错误，
        例如此次空OU——OU=开发部,OU=核心技术部,OU=RAN,OU=上海总部,DC=randolph,DC=com
        的倒数第一、二层都是空OU，但是必须得先删除倒数第一层
        因此在获取所有OU列表的位置get_ous就将获得的结果倒叙(用切片[::-1])
        '''
        res = self.get_ous(attr=['distinguishedName'])
        # 调用ps脚本，防止对象被意外删除×
        modify_right_res = self.del_ou_right(flag=0)
        for i, ou in enumerate(res):
            dn = ou['attributes']['distinguishedName']
            # 判断dd下面是否有用户，没有用户的直接删除
            self.conn.search(search_base=dn, search_filter=USER_SEARCH_FILTER)
            if not self.conn.entries:  # 没有用户存在的空OU，可以进行清理
                try:
                    delete_res = self.conn.delete(dn=dn)
                    if delete_res:
                        logging.info('删除空的OU: ' + dn + ' 成功！')
                    else:
                        logging.error('删除操作处理结果' + str(self.conn.result))
                except Exception as e:
                    logging.error(e)
        else:
            logging.info("没有空OU，OU扫描完成！")
        # 防止对象被意外删除√
        self.del_ou_right(flag=1)

    def disable_users(self, path):
        '''
        @param {type}
        @return:
        @msg: 将AD域内的用户不在csv表格中的定义为离职员工
        '''
        result = ad.handle_excel(path)
        newest_list = []        # 全量员工列表
        for person in result['person_list']:
            job_id, name, dn, email, tel, title, sam, cn = person[0:8]
            dd = str(dn).split(',', 1)[1]
            newest_list.append(name)
        # 查询AD域现有员工
        res = self.get_users(attr=['distinguishedName', 'name', 'cn', 'displayName', 'userAccountControl'])
        for i, ou in enumerate(res):
            ad_user_distinguishedName, ad_user_displayName, ad_user_cn, ad_user_userAccountControl = ou['attributes'][
                'distinguishedName'], ou['attributes']['displayName'], ou['attributes']['cn'], ou['attributes']['userAccountControl']
            rela_dn = "cn=" + str(ad_user_cn)
            # 判断用户不在最新的员工表格中 或者 AD域中某用户为禁用用户
            if ad_user_displayName not in newest_list or ad_user_userAccountControl in DISABLED_USER_FLAG:
                try:
                    # 禁用用户
                    self.conn.modify(dn=ad_user_distinguishedName, changes={'userAccountControl': (2, [546])})
                    logging.info("在AD域中发现不在表格中用户，禁用用户:" + ad_user_distinguishedName)
                    # 移动到离职组 判断OU存在性
                    self.conn.search(DISABLED_BASE_DN, OU_SEARCH_FILTER)    # 判断OU存在性
                    if self.conn.entries == []:                             # 搜不到离职员工OU则需要创建此OU
                        self.create_obj(dn=DISABLED_BASE_DN, type='ou')
                    # 移动到离职组
                    self.conn.modify_dn(dn=ad_user_distinguishedName, relative_dn=rela_dn, new_superior=DISABLED_BASE_DN)
                    logging.info('将禁用用户【' + ad_user_distinguishedName + '】转移到【' + DISABLED_BASE_DN + '】')
                except Exception as e:
                    logging.error(e)

    def create_user_by_excel(self, path):
        '''
        @param path{string} 用于新增用户的表格
        @return:
        @msg:
        '''
        res_dic = self.handle_excel(path)
        for person in res_dic['person_list']:
            user_info = person
            self.create_obj(info=user_info)

    def ad_update(self, path):
        '''AD域的初始化/更新——从表格文件元数据更新AD域:
        判断用户是否在AD域中——不在则新增;
        在则判断该用户各属性是否与表格中相同，有不同则修改;
        完全相同的用户不用作处理;
        '''
        # 准备表格文件
        result = ad.handle_excel(path)
        ori_data = result['person_list']
        try:
            self.del_ou_right(flag=0)       # 防止对象被意外删除×
            with tqdm(iterable=ori_data, ncols=100, total=len(ori_data), desc='处理进度', unit='人') as tqdm_ori_data:    # 封装进度条
                for person in tqdm_ori_data:
                    dn, cn = person[2], person[7]
                    user_info = person
                    dd = str(dn).split(',', 1)[1]
                    # 根据cn判断用户是否已经存在
                    filter_phrase_by_cn = "(&(objectclass=person)(cn=" + cn + "))"
                    search_by_cn = self.conn.search(search_base=ENABLED_BASE_DN, search_filter=filter_phrase_by_cn, attributes=['distinguishedName'])
                    search_by_cn_json_list = json.loads(self.conn.response_to_json())['entries']
                    search_by_cn_res = self.conn.result
                    if search_by_cn == False:                       # 根据cn搜索失败，查无此人则新增
                        self.create_obj(info=user_info)
                    else:
                        old_dn = search_by_cn_json_list[0]['dn']    # 部门改变的用户的现有部门，从表格拼接出来的是新的dn在user_info中带过去修改
                        self.update_obj(old_dn=old_dn, info=user_info)
                    # break                     # 可测试一个例子
                self.del_ou_right(flag=1)       # 防止对象被意外删除√
        except KeyboardInterrupt:
            tqdm_ori_data.close()
            raise
        tqdm_ori_data.close()

    def handle_pwd_expire(self, attr=None):
        '''
        @param {type}
        @return:
        @msg: 处理密码过期 设置密码不过期 需要补全理论和测试
        参考理论地址:
        https://stackoverflow.com/questions/18615958/ldap-pwdlastset-unable-to-change-without-error-showing
        '''
        attr = ['pwdLastSet']
        self.conn.search(search_base=ENABLED_BASE_DN,
                         search_filter=USER_SEARCH_FILTER,
                         attributes=attr)
        result = self.conn.response_to_json()
        res_list = json.loads(result)['entries']
        for l in res_list:
            pwdLastSet, dn = l['attributes']['pwdLastSet'], l['dn']
            modify_res = self.conn.modify(dn, {'pwdLastSet': (2, [-1])})      # pwdLastSet只能给-1 或 0
            if modify_res:
                logging.info('密码不过期-修改用户: ' + dn)

    def update_pwd_file_line(self, old_dn=None, new_dn=None, new_pwd=None):
        '''
        @param dn{string}
        @return: 修改结果
        @msg: 当用户的dn或密码被程序更新，将会在这里更新对应部分的信息
        采用临时文件替换源文件的方式，节省内存，但占硬盘
        参考文章: https://www.cnblogs.com/wuzhengzheng/p/9692368.html
        '''
        with open(PWD_PATH, mode='rt', encoding='utf-8') as file, \
                open('TEMP.txt', mode='wt', encoding='utf-8') as temp_file:
            for line in file:
                if old_dn and new_dn:                   # dn被修改
                    if old_dn in line:
                        line = line.replace(old_dn, new_dn)
                        temp_file.write(line)
                    else:
                        temp_file.write(line)
                elif new_pwd and old_dn:                # 密码被修改
                    if old_dn in line:
                        # 需要正则匹配旧的密码
                        pattern = "PWD: (.+?)\\n"       # 惰性匹配
                        local = re.findall(pattern, line)
                        old_pwd = local[0]
                        line = line.replace(old_pwd, new_pwd)
                        temp_file.write(line)
                    else:
                        temp_file.write(line)
        os.remove(PWD_PATH)
        os.rename('TEMP.txt', PWD_PATH)

    def modify_pwd(self, cn):
        '''
        @param cn{string} 姓名工号 戴东1325
        @return: 修改结果
        @msg: 修改密码
        '''
        # 根据cn判断用户是否已经存在
        filter_phrase_by_cn = "(&(objectclass=person)(cn=" + cn + "))"
        search_by_cn = self.conn.search(search_base=ENABLED_BASE_DN, search_filter=filter_phrase_by_cn, attributes=['distinguishedName'])
        search_by_cn_json_list = json.loads(self.conn.response_to_json())['entries']
        if search_by_cn:
            new_pwd = self.generate_pwd(8)
            old_pwd = ''
            dn = search_by_cn_json_list[0]['dn']
            modify_password_res = self.conn.extend.microsoft.modify_password(dn, new_pwd, old_pwd)
            if modify_password_res:
                logging.info('更新了对象: ' + dn + ' 的密码')
                is_exist = os.path.exists(PWD_PATH)
                if not is_exist:        # 校验密码文件存在性
                    info = 'DN: ' + dn + ' PWD: ' + new_pwd
                    save_res = self.write2txt(PWD_PATH, info)                                       # 将账户密码写入文件中
                    if save_res:
                        logging.info('保存初始化账号密码成功！')
                    else:
                        logging.error('保存初始化账号密码失败: ' + info)
                else:
                    # 若密码修改了需要将密码文件这个人的密码信息更新下
                    with open(PWD_PATH, mode='rt', encoding='utf-8') as file:
                        if dn in file.read():
                            is_exist_pwd_record = True
                        else:
                            is_exist_pwd_record = False
                    if is_exist_pwd_record:     # 若发现此人信息在密码文件里则更新，否则需创建
                        self.update_pwd_file_line(old_dn=dn, new_pwd=new_pwd)
                    else:
                        info = 'DN: ' + dn + ' PWD: ' + new_pwd     # 因为是修改密码，所以dn未修改
                        self.write2txt(PWD_PATH, info)
            else:
                logging.error('更新对象密码失败！: ' + dn)
        else:
            logging.error('查无此人！请检查待修改密码对象格式是否为【姓名工号】')


if __name__ == "__main__":
    # 创建AD域实例
    ad = AD()
    # 修改密码只需要给出 姓名工号 组合的cn     通过√
    # ad.modify_pwd("戴东1325")
    # 同步更新pwd文件     通过√
    # ad.update_pwd_file_line(old_dn='CN=戴东1325,OU=董事会,OU=RAN,OU=上海总部,DC=randolph,DC=com',
    #                         new_dn='CN=戴东1325,OU=RAN,OU=上海总部,DC=randolph,DC=com')
    # 更新AD域     通过√
    # ad.ad_update(RAN_EXCEL)
    # 使用excel新增用户    通过√
    # ad.create_user_by_excel(NEW_RAN_EXCEL)
    # 处理密码过期
    # res_list = ad.handle_pwd_expire()
    # ad.get_ous()
    # 处理源数据    通过√
    # result = ad.handle_excel(TEST_RAN_EXCEL)
    # print(result)
    # 添加OU      通过√
    # ad.create_obj(dn='OU=TEST,DC=randolph,DC=com', type='ou')
    # 分页查询全部user    通过√
    # res = ad.get_users()
    # print(res)
    # 执行powershell命令   通过√
    # ad.del_ou_right(flag=0)
    # 空OU的扫描与删除    通过√
    # ad.scan_ou()
    # 离职员工逻辑    通过√       【M】将禁用员工的处理集成
    # ad.disable_users(RAN_EXCEL)
