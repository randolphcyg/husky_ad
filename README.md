<!--
 * @Author: randolph
 * @Date: 2020-05-27 14:30:31
 * @LastEditors: randolph
 * @LastEditTime: 2020-05-29 16:39:55
 * @version: 1.0
 * @Contact: cyg0504@outlook.com
 * @Descripttion: 
--> 
# husky_ad

### 1.介绍

用python3+ldap3管理windows server2019的AD域; 

实现目标: 

- [x] 批量从excel同步公司组织架构到AD域
- [x] 批量从excel同步公司员工及其信息到AD域
- [x] 批量初始化员工密码
- [x] 提供增删改查简易接口 新增或更新信息接口需要更简易化
- [x] 将重要修改信息与错误信息分离
- [x] 用txt将用户/密码“持久化”
- [x] 使用excel新增用户,只需要将用于新增用户的模板表格更新即可
- [ ] update_obj方法作为基础方法被ad_update方法使用
- [ ] AD域密码过期监控，扫描当前AD域的函数，针对过期时间修改新的密码并覆盖更新
- [ ] 邮件通知模块，根据需要将账号密码的初始化结果/修改结果发送给指定用户

### 2.软件架构
稍后为了更方便使用，会将依赖包的安装改成requirements的，将使用步骤精简化;
这里列出所需要安装和检查的炸药包:

| 项目                                   | 描述                                                       |
| -------------------------------------- | ---------------------------------------------------------- |
| python3.6.8                            | 后端语言                                                   |
| [ldap3](https://ldap3.readthedocs.io/) | 是一个十分优秀且稳健的对active directory域进行管理的炸药包 |
| pandas                                 | 代替python的原生文件读取包,提高处理效率                    |
| winrm                                  | 用来远程连接windows server执行powershell命令               |

### 3.使用说明

1. 将程序及表格模板下载到可以访问AD域的机子上
2. 安装所需要的python依赖库【待优化】
3. 修改AD域配置信息，根据需要准备表格文件，一定包含【工号|姓名|部门】三列，且此三列没有空值(程序中会做校验)，一定是公司拉取的最新的【全量的】数据(表格中没有的员工会做离职处理，禁用账号并移出原组)
4. 可以先调用查询函数测试AD域连通性，然后执行增改等操作【在优化ad_update方法】

### 4.使用举例
#### 4.1.查询AD域测试连接
`ad = AD()`
`ad.get_ous()`
解开main函数中这两句的注释并执行程序;
看到info.log中出现info级别日志即为正常`distinguishedName:CN=Administrator,CN=Users,DC=randolph,DC=com res: True`。

#### 4.2.同步公司名单到AD域
main方法里面将`ad.ad_update(RAN_EXCEL)`解开注释，并运行程序即可;

### 5.问题
欢迎反馈使用中的bug和不便的地方，请提issue，空闲时间将会根据需要进行补充优化~