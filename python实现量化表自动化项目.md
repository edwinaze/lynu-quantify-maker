# python实现量化表自动化项目



## 数据结构



### 学生



- 姓名为索引（key）
- 值为一个字典（dictionary）
  - key : value
  - 学号：number
  - 总分 ：number
  - 加分项名：string
  - 减分项名：string
  - 是否全勤：bool

全勤计算



### 加减分条目



为了做附表, 同时录入信息时也是用这个来录入。

- 项目名：string
- 是否为加/减分项目：bool
- 归属部门：string
- 姓名：string[]
- 分值：number



## 操作分析



### 主界面设计

```
1. 批量文件添加条目（同目录下需有 list.txt）
2. 手动添加单条条目
3. 手动批量添加条目（与1不同是在添加完一条后不会结束）
4. 打印量化表总表
5. 打印量化表附表
6. 打印量化表总表 + 打印量化表附表
```



### 初始化

```
-----------您为第一次使用, 请进行初始化----------
1. 在本程序目录下放入 “班级名单.xlsx” 文件, 第一列中放所有同学的学号, 第二列中放所有同学的姓名,确保该名单顺序和之前做的总表名单顺序一致
2. 重新打开本程序
```

若成功初始化

则直接向 student 读入所有同学名单, 初始总分为2, isattend = true







### 读入数据



假如有这么一条：

```
1 办公室 学生干部 2 何栋宇 潘浩洋 沈涵
```

先读取第一个1, 赋值给 isadd = 1

若为 -1 则返回主界面

接着读取第二个部门 department = "办公室"

项目名 entryname = "学生干部"

加分值 value = 2

后面学生名单则一直读取直到遇见 \n, names[] = {“何栋宇”, “潘浩洋”, “沈  涵”}（若为两个字则加俩空格）

如果是减分则再最后加一条

```
是否为迟到、旷课、早退、晚签到（需扣全勤, 1为是, 0为不是）：
```

notattend = 1/0

#### 然后将数据存入条目列表

```python
entry[entryname] = {
	"isadd" : isadd, 
	"department" ：department,
	"value" : value,
	"name" ; [name]
}
```

#### 接着存入学生列表

```python
for name in names:
	if(isadd)
		student[name]["addentry"].add("entryname", entryname)
		student[name]["total"] += value
	else
		student[name]["subentry"].add("entryname", entryname)
		student[name]["total"] -= value
		if(notattend)  # 被扣分的是该扣全勤的条目
			if(student[name]["isattend"]) # 且还没扣过全勤
				student[name]["total"] -= 2
				student[name]["isattend"] = false
```



### 总表格的新建模板

#### 表头A1:F1合并单元格

内容为 

```
信息技术学院20%year%级软工%class%班 %month%  月份德育量化成绩 综评员：%name% 辅导员签字:________
```

则需输入

```
年级年份 班级 月份 学委名字
21 六 5 何栋宇
```

字体为

```
宋体 14磅
在 软工%class%, %name% 下有下划线
```

#### A2:F2的表头

内容为：

```
学号 姓名 加分项 减分项 总分 学生签名
```

各自的列宽：

```
学号10.85 姓名8.22 加分项63.22 减分项28.89 总分8.89 学生签名11.67
```

字体为：

```
仿宋 14 左右上下居中
```

#### 内容

##### 学号和姓名

读入 "班级档案.xlsx"

学号： 由 A1 读到 max_height, 赋值给 A3 : A(max_helight + 2)

```python
maxheight = backup.max_height
for cells in backup[0][0:maxheight]:
	Sheet1[cells + 3].value = backup[0][cells]
```

姓名同理

##### 加分项

遍历student

```python
cnt = 3 # 从第三行开始
for name, in student:
	# 加分项
	value = "" # 最终输入结果
	if(isattend) value += "全勤+2"
	for entry in name["addentry"]:
		value += " " + entry.first + "+" + entry.second
	Sheet.cell(cnt,3).value = value
	# 减分项
	value = ""
	for entry in name["subentry"]:
		value += entry.first + "+" + entry.second
	Sheet.cell(cnt,4).value = value
	# 总分
	Sheet.cell(cnt,5).value = name["total"]
    cnt++
```



