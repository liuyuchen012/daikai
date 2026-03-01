

# 适用于Windows的多人打卡程序  
### 向中国版权保护中心申请名称 ： 大屏打卡

<!-- PROJECT LOGO -->
<br />

<p align="center">
  <a href="https://github.com/liuyuchen012/daikai">
    <img src="icon.png" alt="Logo" width="80" height="80">
  </a>

  <h3 align="center">"多人打卡系统"</h3>
  <p align="center">
    "使你更好都管理学生"
    <br />
    <a href="https://github.com/liuyuchen012/daikai"><strong>探索本项目的文档 »</strong></a>
    <br />
    <br />
    <a href="https://github.com/liuyuchen012/daikai/">查看Demo</a>
    ·
    <a href="https://github.com/liuyuchen012/daikai/issues">报告Bug</a>
    ·
    <a href="https://github.com/liuyuchen012/daikai/issues">提出新特性</a>
  </p>

</p>

### 应用图片

> V1.3 Offline (Stop Suppost / 停止支持)
> ![ima](01.png)
> ![ima](02.png)

> 

### 作者

刘宇晨 -liuyuchen012- [GitHub](https://github.com/liuyuchen012)

一名生活在天津的初中生

### 项目介绍
本人在班级担任电子教学管理员的职位，受数学老师使用Deepseek制作的打卡程序启发，制作了本程序。

### 程序介绍
#### V1.3 为简易离线版本

你可以使用config.ini配置文件配置你的班级信息。

该配置文件包含如下信息:
```
  [config]
  class_id =      //班级
  nj=             //年级
  z=7             //排版横行
  l=7             //排版竖行
  km=             //科目
  school=         //学校
```


## 如何添加学生信息？
在name.txt中添加学生信息，格式如下：
```txt
1
2
3
```

## 常见错误
-----
### 错误1
#### 错误信息
>Unhandled exception in script
>Failed to execute script 'app' due to unhandled exception:division by zero
>Traceback (most recent call last):
>File "app.py", line 316, in <module>File "app,py",
>line 50,in initFile "app.py",
>line 147, in create _widgetsFile "app,py",
>line 246, in update_statsZeroDivisionError: division by zero
#### 解决办法
>报错是因为你删除了name.txt中的内容并保留了文件
>向文件中添加内容或删除该文件即可解决
-----
### 错误2
#### 错误信息
>Unhandled exception in script
>Failed to execute script 'app' due to unhandled exception:'config'
>Traceback(most recent call last):
>File "app,py", line 26, in <module>File "configparser,py",
>line 979,in getitemKeyError: 'config
-----
### 更多
>如发现更多BUG请向我们[报告](https://github.com/liuyuchen012/daikai/issues)我们会尽力解决


