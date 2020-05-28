<!--
 * @Author: randolph
 * @Date: 2020-05-27 14:30:31
 * @LastEditors: randolph
 * @LastEditTime: 2020-05-29 00:56:23
 * @version: 1.0
 * @Contact: cyg0504@outlook.com
 * @Descripttion: 
--> 
# husky_ad

> 最后更新：2020/5/28

### 1.介绍

用python3+ldap3管理windows server2019的AD域; 

实现目标: 

- [x] 批量从excel同步公司组织架构到AD域
- [x] 批量从excel同步公司员工及其信息到AD域
- [x] 批量初始化员工密码
- [x] 提供增删改查简易接口 新增或更新信息接口需要更简易化
- [x] 将重要修改信息与错误信息分离
- [x] 用txt将用户/密码“持久化”
- [ ] AD域密码过期监控
- [ ] update_obj方法需要优化为根据name自动判断dn并更新
- [ ] ad_update方法需要增加健壮性检验

### 2.软件架构

这里列出所需要安装和检查的炸药包:

| 项目                                   | 描述                                                       |
| -------------------------------------- | ---------------------------------------------------------- |
| python3.6.8                            | 后端语言                                                   |
| [ldap3](https://ldap3.readthedocs.io/) | 是一个十分优秀且稳健的对active directory域进行管理的炸药包 |
| pandas                                 | 代替python的原生文件读取包,提高处理效率                    |
| winrm                                  | 用来远程连接windows server执行powershell命令               |

### 3.使用说明

将脚本下载到可以访问AD域的机子上，批量操作需要准备全量的公司人员清单，需要修改程序开头配置信息并进行测试。

1. 修改ad域的配置信息并进行测试连通性
2. 将excel表格拷贝到代码同目录下并修改配置信息;
    表格的要求:
    1. 一定包含【工号|姓名|部门】三列，且此三列没有空值(程序中会做校验)
    2. 每次一定是公司拉取的最新的【全量的】数据(表格中没有的员工会做离职处理，禁用账号并移出原组)
3. 执行“面对过程”的函数即可，检查日志，确定没有问题

### 3.AD域即将过期怎么办

我会写一个批量扫描当前AD域的函数，针对过期时间修改新的密码并覆盖更新