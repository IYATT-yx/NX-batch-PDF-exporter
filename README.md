# NX-batch-PDF-exporter【Siemens NX 批量 PDF 图纸导出工具】

2025/11/30  
本工具是用于批量导出 Siemens NX 图纸到 PDF 文件的，可以多选图纸文件，工具依次导出 PDF。  
本来计划是上周写好的，当时初步拟定了框架，结果上周五、周六去出差了，回来也不想写了。这周我又沉迷于看动漫去了，下班时间都花完了，今天星期天又想起来，现在晚上十点多点，基本功能应该是可以用了。我也不知道会不会有隐藏的 bug，只有后面使用中测试了，有问题就再改。  
![alt text](doc/image1.png)  

## 测试环境

* Siemens NX 2506
* 内置 Python 版本：3.12.8  

## 使用方法

### 方式一

按`Alt`+`F8`打开“操作记录管理器”，浏览本工具中“NX-batch-PDF-exporter.py”所在路径，点击运行即可  
![alt text](doc/image2.png)

### 方式二

在菜单栏空白处右键，点击“定制”  
![alt text](doc/image3.png)  

在命令选项卡下，左侧点新建项，右边点击新建用户命令拖动到菜单栏上  
![alt text](doc/image4.png)  

在新建的图标上右键，可以设置名称，再点击最下面编辑操作  
![alt text](doc/image5.png)  

浏览 “NX-batch-PDF-exporter.py” 所在路径即可  
![alt text](doc/image6.png)  

现在可以通过直接点击图标运行了  
![alt text](doc/image7.png)