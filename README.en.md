<!--
 * @Author: randolph
 * @Date: 2020-05-27 14:30:31
 * @LastEditors: randolph
 * @LastEditTime: 2020-05-27 15:16:35
 * @version: 1.0
 * @Contact: cyg0504@outlook.com
 * @Descripttion: 
--> 
# husky_ad

> Last update: 2020/5/27

### 1 Introduction

Use python3 + ldap3 to manage the AD domain of windows server2019;

Goals: 

-[x] Synchronize company organization structure from excel to AD domain in batches
-[x] Batch sync company employees and their information from excel to AD domain
-[x] Batch initialize employee passwords
-[x] Provide simple interface for adding, deleting, modifying and checking
-[] ...

### 2. Software Architecture

The explosive packages that need to be installed and checked are listed here:

| Item | Description |
| -------------------------------------- | ---------- ------------------------------------------------ |
| python3.6.8 | Backend language |
| [ldap3] (https://ldap3.readthedocs.io/) | is an excellent and robust explosive package for managing the active directory domain |
| pandas | Instead of Python's native file reading package, improve processing efficiency |
| winrm | Used to remotely connect to the windows server to execute powershell commands |

### 3. Instructions

1. Modify the configuration information of ad domain and test connectivity
2. Copy the excel form to the same directory as the code and modify the configuration information
    Requirements for the form:
    1. It must contain three columns of [work number | name | department], and there are no empty values ​​in these three columns (verification will be done in the program)
    2. Each time it must be the latest [full amount] data pulled by the company (employees who are not in the form will do the resignation process, disable the account and move out of the original group)
3. Just execute the "face to process" function, check the log and make sure there is no problem