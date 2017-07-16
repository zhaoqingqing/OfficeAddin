

### 运行和测试环境

我的运行和测试环境：

操作系统：windows 10 x64 企业版/专业版，建议Windows 7及以上操作系统。

办公软件：WPS 10个人版、Office365

## Office API

wps也可以调用实现了ms office  api 接口的类库

### Ribbon(界面)

### COM

## 源代码二次开发？

此仓库仅仅是个示例，如果你需要扩展更多功能，欢迎Fork此仓库进行修改。

开发工具：建议使用visual studio 2010及以上版本，我使用的是visual studio 2017，安装**VSTO(visual studio tool for office)**

项目类型：Visual C# - Office - SharePoint  - **Excel 2013和2016 VSTO外接程序**

### 生成遇到错误？

如果在生成过程中遇到以下错误：

``` shell
无法注册程序集“D:\Git\OfficeAddin\ExcelAddIn\bin\Debug\ExcelAddIn.dll”- 拒绝访问。请确保您正在以管理员身份运行应用程序。对注册表项“HKEY_CLASSES_ROOT\ExcelAddIn.QuickRibbon”的访问被拒绝。	ExcelAddIn		
```
确保代码正确生成，将ExcelAddIn.dll 添加到Excel的COM加载项中。