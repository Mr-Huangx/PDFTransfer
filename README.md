> 本项目使用pyqt开发了一款word模板编写和转换的功能。



## 1、简介

本项目使用pyqt开发了一款word模板编写和转换的功能。实现了以下功能：

1. 使用docx库，完成docx文件的读取，并提取其中段落和表格的占位符（占位符格式``{{xx}}``）
   根据提取的占位符，动态生成表单控件。
2. 使用docx2pdf库，完成docx文件模板的填充，并转换为对应的pdf。
3. 使用pandas库，完成excel报价表单的生成。并允许手动修改表单的字段信息。
4. 使用logging库，完成程序的不同级别日志输出。
5. 使用pyinstaller完成程序的打包功能



## 2、本项目运行示意图：

### 2.1 帮助页面

![help_page](.\imgs\help_page.png)



### 2.2 模板填写表单页

![main_ui](.\imgs\main_ui.png)



### 2.3 excel表单填写

![excel_ui](.\imgs\excel_ui.png)



## 3、环境

- OS：windows11

- python：3.10.9



#### 3.1 安装依赖

```shell
pip install -r requirements.txt
```

#### 3.2 运行程序

​	第一步，进入MainUI

​	第二步，运行

```shell
cd MainUI
python main.py
```



### 3.3 打包exe

​	安装pyinstaller

```
pip install pyinstaller
```

​	打包

```
pyinstaller main.spec
```

