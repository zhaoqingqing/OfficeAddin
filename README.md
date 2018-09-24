## OfficeAddin

本文可能会用到以下简写

ms office是microsoft office的简写

vs是visual studio 的简写

wps是金山wps办公软件的简写

### Office(办公软件)扩展

扩展office办公软件的功能，目前示例 扩展excel的Ribbon界面

![https://github.com/zhaoqingqing/OfficeAddin/blob/master/Docs/wps_ribbon.png](https://github.com/zhaoqingqing/OfficeAddin/blob/master/Docs/wps_ribbon.png)

### 运行和测试环境

我的运行和测试环境：

操作系统：windows 10 x64 企业版/专业版，建议Windows 7及以上操作系统。

办公软件：WPS 10个人版、Office365

其它条件：需要安装.NET Framwork 4.0 ，Win7(自带.Net3.5) 如更新过Windows补丁则会更新到.Net4.0 ， WIN8/10用户无需额外安装。

### 使用说明

1. 使用vs生成代码，编译出 ExcelAddIn.dll
2. 双击 **将ExcelAddin添加到注册表.reg**  同时为ms office和wps添加注册项
3. 如果是wps，在菜单栏选择 **开发工具** - **COM加载项** - 添加 **ExcelAddIn.dll** 为加载项 - 点击 **确定**，在菜单栏的ribbon即可看到

如果你不想使用vs编译代码，可以使用我编译好的dll(编译环境为.net 4.0) [ExcelAddin.dll](https://github.com/zhaoqingqing/OfficeAddin/blob/master/ExcelAddIn/release_bin/ExcelAddIn.dll)


### 卸载已安装的加载项

理论上在 **开发工具** - **COM加载项** 删除不需要的加载项。

如果遇到问题，wps请在安装目录下找到ksomisc.exe，双击进行修复。

office请通过更改进行修复。

## Microsoft Office API

经测试wps可以通过COM加载：使用ms office  api 类库编写的接口，所以理论上同一套代码可同时运行于微软的Office办公软件和金山WPS之上。

### Ribbon(界面)



### COM

TODO 知识补充



## 源代码二次开发？

此仓库仅仅是示例，如果你需要扩展更多功能，欢迎Fork此仓库进行修改。

开发工具：建议使用**visual studio 2010及以上版本**，我使用的是visual studio 2017，安装**VSTO(visual studio tool for office)**

项目类型：使用vs新建：Visual C# - Office - SharePoint  - **Excel 2013和2016 VSTO外接程序**

### 生成遇到错误？

如果在生成过程中遇到以下错误：

``` shell
无法注册程序集“D:\Git\OfficeAddin\ExcelAddIn\bin\Debug\ExcelAddIn.dll”- 拒绝访问。请确保您正在以管理员身份运行应用程序。对注册表项“HKEY_CLASSES_ROOT\ExcelAddIn.QuickRibbon”的访问被拒绝。	ExcelAddIn		
```
确保代码正确生成，将ExcelAddIn.dll 添加到Excel的COM加载项中。



## TODO

- [x] 获取当前Excel的文件名 
- [x] 保存当前编辑的Excel

