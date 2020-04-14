# Tencent-classroom-attendance-assistant
 腾讯课堂考勤助手：一款自动化考勤数据处理工具
 
腾讯课堂导出的考勤数据只有上课学生数据，使得老师们需要对比学生名单逐个看学生有没有上课，或者是上了多久课，找出缺勤的学生，考勤起来需要耗费不少的时间。所以这几天用了空闲时间写了一款处理考勤数据的工具，这个工具能够自动处理腾讯课堂导出的考勤数据，生成本班学生每个科目缺勤次数，减少老师的工作量。

[**软件下载**][1]（目前只打包了exe文件）

本程序基于python3开发，不得不说python处理excel真多坑，一开始使用了xlrd、xlwt、xlutils三个库，发现xlwt居然不能保存为xlsx？？？读写分开两个库也非常麻烦。后来又发现了一个更好的库openpyxl单个库同时支持读写，所以写到后面又换成了这个库。本程序没有gui！！！只要一个丑的一批的控制台，为什么呢？很简单小白不会写gui TAT。

版本说明
----
**V1.0.0**

支持导出本班学生每个科目缺勤次数
已知问题：输出全级数据时，如果缺少该班本节课考勤表时，会全班记为缺勤
（目前仅能输出单个班数据，将会在下个版本修复）

使用教程
----
![1.png][2]
![2.png][3]
![3.png][4]
![4.png][5]
![5.jpg][6]
![6.jpg][7]
![7.png][8]

配置文件
----
**使用demo格式无需修改**

    {
        "name_x": 1, #学生名单中开始读取学生姓名的行（值需要减1）
        "name_y": 3, #学生名单中开始读取学生姓名的列（值需要减1）
        "txkt_start_x": 5, #考勤表中开始读取学生数据的行（值需要减1）
        "txkt_duration_y": 7, #考勤表中开始读取学生上课时间的列（值需要减1）
        "txkt_name_y": 3, #考勤表中开始读取学生姓名的列（值需要减1）
        "class_y": 4, #学生名单中学生数据截至的列，将会在列加1写入考勤数据
        "data_path": "./data", #考勤表路径
        "class_path": "./class.xlsx", #学生名单路径
        "min_class": 20, #最少上课时间，少于将会记为缺勤
        "min_num": 10 #该节课最少有效上课学生，如设置过小，其它班进错科室会导致本班学生缺勤
    }

  [1]: https://github.com/Xiaobin2333/Tencent-classroom-attendance-assistant/releases
  [2]: https://search.pstatic.net/common?type=origin&src=https://www.mrchung.cn/usr/uploads/2020/04/168203530.png
  [3]: https://search.pstatic.net/common?type=origin&src=https://www.mrchung.cn/usr/uploads/2020/04/2897639418.png
  [4]: https://search.pstatic.net/common?type=origin&src=https://www.mrchung.cn/usr/uploads/2020/04/2006752868.png
  [5]: https://search.pstatic.net/common?type=origin&src=https://www.mrchung.cn/usr/uploads/2020/04/4291132881.png
  [6]: https://search.pstatic.net/common?type=origin&src=https://www.mrchung.cn/usr/uploads/2020/04/126812428.jpg
  [7]: https://search.pstatic.net/common?type=origin&src=https://www.mrchung.cn/usr/uploads/2020/04/2036134815.jpg
  [8]: https://search.pstatic.net/common?type=origin&src=https://www.mrchung.cn/usr/uploads/2020/04/745169198.png
