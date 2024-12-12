# 开发说明
## 安装依赖
- python>=3.10.12
- pip install -r requirements.txt

## 打包说明
- Pyinstaller -F SameTopicSearch.py
- 打好会生成一个dist文件夹，它下边就是生成了exe程序

***
# 使用说明
## 参数介绍
- 第一个参数（必填）：用来校验的目录地址,需注意：1、windows盘符需要大写；2、路径中如果有空格需要用引号包裹起来
- 第二个参数（必填）：要校验的excel文件名称，需要将校验文件放在执行程序同级目录
- 第三个参数（必填）：要校验的字段（列）：[topic:题干；option:选项A]；
- 第四个参数: 是否拆词校验：[0:不拆（精准搜索）；1:拆，不填，默认为0]
- 举例：SameTopicSearch "F:\\资料库" test.xls topic 0