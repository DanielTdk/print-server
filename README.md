# Print-Server 服务

## 服务介绍
本项目使用golang开发，通过go-ole 调用Windows系统的COM接口和方法(office、bartender)

这里直接提供exe二进制可执行文件，开箱即用

仅支持windows服务器

1、office excel 打印excel文件服务需要本地安装建议office 16 以上版本；

   实现方式-调用office com组件调起打印功能；

2、bartender 打印标签服务需要本地安装Bartender软件

   实现方式-调用bartender com组件调起打印功能


启动方式：


```bash
print-server.exe >> 日志.log
```

## 服务地址

 `http://localhost:8080`
## 1、接口 excel打印接口：
```bash
http://localhost:8080/v1/api/excel/print
Post form-data multipart/form-data; 
"file" :/path/to/file (multipart) -- 文件上传
"printConfig": "" -- 字符串
示例数据:
"{ 
  "marginConfig": { //打印设置 页边距
    "marginTop": 54.1485,
    "marginBottom": 54.1485,
    "marginLeft": 18.144,
    "marginRight": 18.144,
    "marginFooter": 21.546,
    "marginHeader": 21.546
  },
  "printerName": "打印机名称",
  "copies": 1, //份数
  "paperSize": 9 // A3是8，A4是9，A5是11
}"


```
## 2、接口 bartender打印接口：
```bash

http://localhost:8080/v1/api/bartender/print
Post  application/json
示例数据:
{
    "btwName": "bartender模板.btw",// 模板文件名
    "dataType": 1, //打印方式 1 文本数据源  2 嵌入式数据源 
    "path": "D:/bartender_template", //模板本地路径地址
    "embeddedDataList":  //嵌入式数据源数据项列表
        [
            {
                "name": "key1",
                "value": "key1"
            },
            {
                "name": "key2",
                "value": "key2"
            },
            {
                "name": "key3",
                "value": "key3"
            }
    ],
    "txtName": "new1.dat", //文本数据源需要
    "title": "test title", //文本数据源需要
    "dataList": ["key1","key2","key3"], // 文本数据源需要
    "identicalCopiesOfLabel": 1, //设置打印份数
    "numberSerializedLabels": 1, //设置自增序列
    "printerName": "Microsoft Print to PDF" //填写本地打印机名称
}


```
