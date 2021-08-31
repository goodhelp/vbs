# vbs代码

## 介绍
vbs常用函数代码，包括一些特色函数，杀毒软件可能会报毒，使用notepad++等软件可直接查看源代码，未进行任何加密
最大的特色，可以让vbs使用windows api函数，api能办到的功能，vbs也能做到!
自定义vbs类的函数，见demo\demo.vbs文件

## 软件架构
lib目录的MyVbsClass.vbs和dynwrapx.dll为核心文件，即vbs自定义类
demo为示例代码，
维护通道是使用vbs自定义类写的网吧开机维护通道


## 安装教程

1.  把lib目录复制一份
2.  新建一目录，引用lib中的MyVbsClass.vbs即可使用其全部函数，引用方法见demo\demo.vbs

