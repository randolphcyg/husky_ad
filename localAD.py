#!/usr/bin/env python
# -*- encoding: utf-8 -*-
'''
@Author: randolph
@Date: 2020-05-05 15:48:26
@LastEditors: randolph
@LastEditTime: 2020-05-27 16:35:34
@version: 2.0
@Contact: cyg0504@outlook.com
@Descripttion: 优化了日志的中文编码、winrm的操作、随机密码生成逻辑、AD域查询改成分页以适应超过1000的查询情况
表格文件的解析稳健性——增加对列数的判断、对每列对应何种属性的自动判断；
需要将AD域的扫描流程整合成一键操作的；每一步骤都设置提示，让用户使用的时候出现错误有事务的回退；
'''
# 基础服务
import json
import winrm
import string
import random
import logging
# 数据处理
import pandas as pd
# LDAP3
from ldap3 import Server, Connection, SIMPLE, SYNC, ALL, SASL, NTLM, ALL_ATTRIBUTES, MODIFY_REPLACE, SUBTREE
# 日志设置
LOG_FORMAT = "%(asctime)s  %(levelname)s  %(filename)s  %(lineno)d  %(message)s"
LOG_FILE = open("localAD.log", encoding="utf-8", mode="a")
logging.basicConfig(stream=LOG_FILE, format=LOG_FORMAT, level=logging.INFO)
# AD域设置
LDAP_IP = '192.168.255.223'                                 # LDAP本地服务器IP
USER = 'CN=Administrator,CN=Users,DC=randolph,DC=com'       # LDAP本地服务器IP
PASSWORD = "QQqq#123"                                       # LDAP本地服务器管理员密码
# excel表格
BILIBILI_EXCEL = "ran_list.xlsx"                        # 原始造的数据
TEST_BILIBILI_EXCEL = "test_ran_list.xlsx"              # 测试用表格
# WINRM信息 无需设置
WINRM_USER = 'Administrator'
WINRM_PWD = PASSWORD


class AD(object):
    '''AD域的操作
    '''

    def __init__(self):
        '''初始化 AD域连接
        '''
        SERVER = Server(host=LDAP_IP,
                        port=636,                                       # 636安全端口
                        use_ssl=True,
                        get_info=ALL,
                        connect_timeout=3)                              # 连接超时为3秒
        self.conn = Connection(
            server=SERVER,
            user=USER,
            password=PASSWORD,
            auto_bind=True,
            read_only=False,                                            # 禁止修改数据True
            receive_timeout=3)                                          # 3秒内没返回消息则触发超时异常

        self.disabled_base_dn = 'OU=resigned,DC=randolph,DC=com'        # 离职账户所在OU
        self.enabled_base_dn = 'OU=上海总部,DC=randolph,DC=com'         # 正式员工账户所在OU
        self.user_search_filter = '(objectclass=user)'                  # 只获取用户对象
        self.ou_search_filter = '(objectclass=organizationalUnit)'      # 只获取OU对象
        self.disabled_user_flag = [514, 546, 66050, 66080, 66082]       # 禁用账户
        self.enabled_user_flag = [512, 544, 66048, 262656]              # 启用账户

    def check_conn(self):
        try:
            logging.info("username:%s res: %s" % (USER, self.conn.bind()))
            return self.conn
        except BaseException:
            logging.warning("username:%s res: %s" % (USER, self.conn.bind()))
            return False
        finally:
            self.conn.closed

    def get_users(self, attr=ALL_ATTRIBUTES):
        '''
        @param {type}
        @return: total_entries所有用户
        @msg: 获取所有用户
        '''
        entry_list = self.conn.extend.standard.paged_search(
            search_base=self.enabled_base_dn,
            search_filter=self.user_search_filter,
            search_scope=SUBTREE,
            attributes=attr,
            paged_size=5,
            generator=False)                                        # 关闭生成器，结果为列表
        total_entries = 0
        for entry in entry_list:
            total_entries += 1
            # print(entry['attributes']['distinguishedName'])
            # for k, v in zip(entry['attributes'].keys(), entry['attributes'].values()):
            #     print(k, v)
            # break
        # print(entry['dn'], entry['attributes'])
        # print('共查询到记录条目:', total_entries)
        return entry_list

    def get_ous(self, attr=None):
        '''
        @param {type}
        @return: res所有OU
        @msg: 获取所有OU
        '''
        self.conn.search(search_base=self.enabled_base_dn,
                         search_filter=self.ou_search_filter,
                         attributes=attr)
        res = self.conn.response_to_json()
        res = json.loads(res)['entries']
        return res

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
                    row[2] = 'CN=' + str(name + str(job_id)) + ',' + 'OU=' + ',OU='.join(row[2].split('.')[::-1]) + ',' + self.enabled_base_dn
                    row.append('RAN' + str(job_id).zfill(6))        # 增加登录名列,对应AD域user的 sAMAccountname 属性
                    row.append(name + str(job_id))                  # 增加CN列,对应user的 cn 属性
                # 3.开始处理返回字典
                result = dict()                         # 返回字典
                if a > 1000:
                    result['page_flag'] = True
                else:
                    result['page_flag'] = False
                result['person_list'] = person_list
                return result
        except Exception as e:
            logging.error(e)

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
        pwd = ''.join(pwd_list)
        return pwd

    def del_ou_right(self, flag):
        '''
        @param cmd_l{list} 待执行的powershell命令列表
        @return: True/False
        @msg: 连接远程windows并批量执行powershell命令
        '''
        # powershell命令，用于打开/关闭OU是否被删除的权限
        enable_del = ["Import-Module ActiveDirectory",
                      "Get-ADOrganizationalUnit -filter * -Properties ProtectedFromAccidentalDeletion | where {"
                      "$_.ProtectedFromAccidentalDeletion -eq $true} |Set-ADOrganizationalUnit "
                      "-ProtectedFromAccidentalDeletion $false"]
        disable_del = ["Import-Module ActiveDirectory",
                       "Get-ADOrganizationalUnit -filter * -Properties ProtectedFromAccidentalDeletion | where {"
                       "$_.ProtectedFromAccidentalDeletion -eq $false} |Set-ADOrganizationalUnit "
                       "-ProtectedFromAccidentalDeletion $true"]
        flag_map = {0: enable_del, 1: disable_del}

        try:
            win = winrm.Session('http://' + LDAP_IP + ':5985/wsman', auth=(WINRM_USER, WINRM_PWD))
            for cmd in flag_map[flag]:
                ret = win.run_ps(cmd)
            if ret.status_code == 0:      # 调用成功
                if flag == 0:
                    logging.info("防止对象被意外删除×")
                elif flag == 1:
                    logging.info("防止对象被意外删除√")
                return True
            else:
                return False
        except Exception as e:
            return False
            logging.error(e)

    def create_obj(self, dn, type, attr=None):
        '''
        @param dn{string}, type{string}user/ou
        @return: res新建结果, self.conn.result修改结果
        @msg:增加对象
        '''
        object_class = {'user': ['user', 'posixGroup', 'top'],
                        'ou': ['organizationalUnit', 'posixGroup', 'top'],
                        }
        res = self.conn.add(dn=dn, object_class=object_class[type], attributes=attr)
        if type == 'user':                                                                  # 如果是用户时
            new_pwd = self.generate_pwd(8)
            old_pwd = ''
            self.conn.extend.microsoft.modify_password(dn, new_pwd, old_pwd)                # 初始化密码
            self.conn.modify(dn, {'userAccountControl': [('MODIFY_REPLACE', 512)]})         # 激活用户
            logging.info('dn:【' + str(dn) + '】' + 'pwd:【' + str(new_pwd) + '】')         # 记录密码修改
            self.conn.modify(dn, {'pwdLastSet': (2, [0])})                                  # 设置第一次登录必须修改密码
        return res, self.conn.result

    def del_obj(self, dn):
        '''
        @param dn{string}
        @return: res修改结果
        @msg: 删除对象
        '''
        res = self.conn.delete(dn=dn)
        return res

    def update_obj(self, dn, attr=None):
        '''
        @param {type}
        @return:
        @msg: 更新对象，这个还不清楚
        '''
        changes_dic = {}
        for k, v in attr.items():
            if not self.conn.compare(dn=dn, attribute=k, value=v):
                if k == "name":
                    res = self.__rename_obj(dn=dn, newname=attr['name'])     # 改过名字后，DN就变了,这里调用重命名的方法
                    if res['description'] == "success":
                        if "CN" == dn[:2]:
                            dn = "cn=%s,%s" % (attr["name"], dn.split(",", 1)[1])
                        if "OU" == dn[:2]:
                            dn = "DN=%s,%s" % (attr["name"], dn.split(",", 1)[1])
                if k == "DistinguishedName":                            # 如果属性里有“DistinguishedName”，表示需要移动User or OU
                    self.__move_obj(dn=dn, new_dn=v)                  # 调用移动User or OU 的方法
                changes_dic.update({k: [(MODIFY_REPLACE, [v])]})
                self.conn.modify(dn=dn, changes=changes_dic)
        return self.conn.result

    def __rename_obj(self, dn, newname):
        '''
        @param newname{type}新的名字，User格式："cn=新名字";OU格式："OU=新名字"
        @return: 修改结果
        @msg: 重命名对象
        '''
        self.conn.modify_dn(dn, newname)
        return self.conn.result

    def __move_obj(self, dn, new_dn):
        '''
        @param {type}
        @return:
        @msg: 移动对象到新OU
        '''
        relative_dn, superou = new_dn.split(",", 1)
        res = self.conn.modify_dn(dn=dn, relative_dn=relative_dn, new_superior=superou)
        return res

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
        self.conn.search(ou, self.ou_search_filter)  # 判断OU存在性

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
        res = self.get_ous(attr=['distinguishedName'])
        for i, ou in enumerate(res):
            dn = ou['attributes']['distinguishedName']
            # 判断dd下面是否有用户，没有用户的直接删除
            self.conn.search(search_base=dn, search_filter=self.user_search_filter)
            if not self.conn.entries:  # 没有用户存在的空OU，可以进行清理
                try:
                    # 调用ps脚本，防止对象被意外删除×
                    modify_right_res = self.del_ou_right(flag=0)
                    if modify_right_res:
                        self.conn.delete(dn=dn)
                    if self.conn.result['result'] == 0:
                        logging.info('删除空的OU: ' + dn + ' 成功！')
                    else:
                        logging.error('删除操作处理结果' + str(self.conn.result))
                    # 防止对象被意外删除√
                    self.del_ou_right(flag=1)
                except Exception as e:
                    logging.error(e)
        else:
            logging.info("没有空OU，OU扫描完成！")

    def disable_user(self):
        '''
        @param {type}
        @return:
        @msg: 将AD域内的用户不在csv表格中的定义为离职员工
        '''
        result = ad.handle_excel(TEST_BILIBILI_EXCEL)
        newest_list = []        # 全量员工列表
        for person in result['person_list']:
            job_id, name, dn, email, tel, title, sam, cn = person[0:8]
            # print(job_id, name, dn, email, tel, title, sam, cn)
            dd = str(dn).split(',', 1)[1]
            newest_list.append(name)
        # 查询AD域现有员工
        res = self.get_users(attr=['distinguishedName', 'name', 'cn', 'displayName', 'userAccountControl'])
        for i, ou in enumerate(res):
            ad_user_distinguishedName, ad_user_displayName, ad_user_cn, ad_user_userAccountControl = ou['attributes'][
                'distinguishedName'], ou['attributes']['displayName'], ou['attributes']['cn'], ou['attributes']['userAccountControl']
            rela_dn = "cn=" + str(ad_user_cn)
            print(ad_user_distinguishedName, ad_user_displayName, ad_user_cn, ad_user_userAccountControl, rela_dn)
            # 判断用户不在最新的员工表格中 或者 AD域中某用户为禁用用户
            if ad_user_displayName not in newest_list or ad_user_userAccountControl in self.disabled_user_flag:
                try:
                    # 禁用用户
                    self.conn.modify(dn=ad_user_distinguishedName, changes={'userAccountControl': (2, [546])})
                    logging.info("禁用用户:" + ad_user_distinguishedName)
                    # 移动到离职组 判断OU存在性
                    self.conn.search(self.disabled_base_dn, self.ou_search_filter)  # 判断OU存在性
                    if self.conn.entries == []:         # 搜不到离职员工OU则需要创建此OU
                        self.create_obj(dn=self.disabled_base_dn, type='ou')
                    # 移动到离职组
                    self.conn.modify_dn(dn=ad_user_distinguishedName, relative_dn=rela_dn, new_superior=self.disabled_base_dn)
                    logging.info('将禁用用户【' + ad_user_cn + '】转移到【' + self.disabled_base_dn + '】')
                except Exception as e:
                    logging.error(e)

    def ad_update(self):
        '''ad域的初始化或更新: 将从表格处理好的数据同步到AD域：
        如果AD域没有OU，则创建OU；
        如果没有人则创建；
        '''
        result = ad.handle_excel(TEST_BILIBILI_EXCEL)
        # print(result['page_flag'])
        for person in result['person_list']:
            job_id, name, dn, email, tel, title, sam, cn = person[0:8]
            print(job_id, name, dn, email, tel, title, sam, cn)
            dd = str(dn).split(',', 1)[1]
            # 通过表格中的路径去搜索AD域中对应的用户，如果能搜到说明没改变，略过；
            # 如果没搜到，有可能是该用户调整了位置|或者该用户是新用户，没有创建
            self.conn.search(dn, '(objectclass=user)', attributes=['distinguishedName'])
            if self.conn.result['result'] == 0:      # 未发生变化的用户
                pass
            else:
                filter_phrase = "(&(objectclass=person)(cn=" + cn + "))"
                self.conn.search(search_base=self.disabled_base_dn, search_filter=filter_phrase, attributes=['*'])
                entry = self.conn.entries
                if entry:
                    rela_dn = "cn=" + str(cn)
                    # print("待修改用户 " + str(entry[0].distinguishedName), rela_dn, dd)
                    try:
                        self.conn.modify_dn(dn=entry[0].distinguishedName, relative_dn=rela_dn, new_superior=dd)
                        if self.conn.result['result'] == 0:
                            logging.info("modify_dn " + str(entry[0].distinguishedName), rela_dn, dd)
                        else:
                            if self.check_ou(dd):
                                self.conn.modify_dn(dn=str(entry[0].distinguishedName), relative_dn=str(rela_dn),
                                                    new_superior=str(dd))
                                logging.info("modify_dn " + str(entry[0].distinguishedName), rela_dn, dd)
                    except Exception as e:
                        logging.error(e)
                else:  # 需要新增user
                    if self.check_ou(dd):
                        user_attr = {'sAMAccountname': sam,      # 登录名
                                     'userAccountControl': 544,  # 启用账户
                                     'title': title,             # 头衔
                                     'givenName': name[0:1],     # 姓
                                     'sn': name[1:],             # 名
                                     'displayname': name,        # 姓名
                                     'mail': email,              # 邮箱
                                     'telephoneNumber': tel,     # 电话号
                                     }
                        self.create_obj(dn=dn, type='user', attr=user_attr)
            # break


if __name__ == "__main__":
    # 0.创建一个实例
    ad = AD()
    # 1.检测AD域连通性 # 通过√
    ad.check_conn()
    # 2.处理源数据    通过√
    # result = ad.handle_excel(TEST_BILIBILI_EXCEL)
    # print(result)
    # 3.添加人员    通过√
    # ad.create_obj('CN=王大锤1,OU=测试,DC=randolph,DC=com', 'user', attr={
    #             'sAMAccountname': 'RAN000001',      # 登录名
    #             'userAccountControl': 544,  # 启用账户
    #             'title': '技术顾问',             # 头衔
    #             'givenName': "王",     # 姓
    #             'sn': "大锤",             # 名
    #             'displayname': "王大锤",        # 姓名
    #             'mail': "dachui.wang@ran-china.com",              # 邮箱
    #             'telephoneNumber': 1502510654,     # 电话号
    #         })
    # 4.添加OU      通过√
    # ad.create_obj('OU=RAN,DC=randolph,DC=com', 'ou')
    # 5.删除对象    通过√
    # ad.del_obj('OU=RAN,DC=randolph,DC=com')
    # 6.分页查询全部user    通过√
    # ad.get_users()
    # 7.更新AD域     通过√ 【对于新增的没有问题】  @@@@@修改的待修改@@@@@
    # ad.ad_update()
    # 8.执行powershell命令   通过√
    # ad.del_ou_right(flag=0)
    # 9.空OU的扫描与删除    通过√
    # ad.scan_ou()
    # 10.离职员工逻辑    通过√       【M】将禁用员工的处理集成
    # ad.disable_user()
