> VBS



---

# 目录

---

* [1.轻松开发脚本](#1.轻松开发脚本)
* [2.VBScript 入门](#2.VBScript 入门)
* [3.数据类型](#3.数据类型)
* [4.变量和程序](#4.变量和程序)



---

# 1.轻松开发脚本



## 1.1  写我们第一个脚本

```vbscript
MsgBox "Hello World"
```

## 1.2 编辑脚本内容

```vbscript
option Explicit

dim ask
dim title
dim answer
ask = "Hey,what's your name?"
title = "Who are you?"
answer = InputBox(ask,title)
if answer = vbEmpty then
	MsgBox "You want to cancel? Fine!"	
	WSCript.Quit
elseif answer = "" then
	MsgBox  "You want to stay incongnito,all right with me..."	
else
	MsgBox  "Welcome " & answer & "!"
end if
```



# 2.VBScript 入门



# 3.数据类型

> 变量可以保存不同类型的数据：数字、日期、文本和其他更专门或更复杂的类别。这些可以划分值的类别称为数据类型。



| 子类型       | 描述                                                         |
| ------------ | ------------------------------------------------------------ |
| `Empty`      | 未初始化的**Variant**。对于数值变量，值为0；对于字符串变量，值为零长度字符串 ("")。 |
| `NULL`       | 不包含任何有效数据的Variant                                  |
| `Boolean`    | 包含True或False。                                            |
| `Byte`       | 包含0到255之间的整数。                                       |
| `Integer`    | 包含-32,768到32,767之间的整数。                              |
| `Currency`   | -922,337,203,685,477.5808到922,337,203,685,477.5807。        |
| `Long`       | 包含-2,147,483,648到2,147,483,647之间的整数。                |
| `Single`     | 包含单精度浮点数，负数范围从-3.402823E38到-1.401298E-45，正数范围从1.401298E-45到3.402823E38。 |
| `Double`     | 包含双精度浮点数，负数范围从-1.79769313486232E308到-4.94065645841247E-324，正数范围从4.94065645841247E-324到1.79769313486232E308。 |
| `Date(Time)` | 包含表示日期的数字，日期范围从公元100年1月1日到公元9999年12月31日。 |
| `String`     | 包含变长字符串，最大长度为20亿个字符                         |
| `Object`     | 包含对象                                                     |
| `Error`      | 包含错误号                                                   |







```vbscript
option Explicit

dim a 
dim b
dim c
dim d
dim e
dim f

a = "hello world"
b = 12
c = 12.3
d = false
f =#11/22/2020#

MsgBox TypeName(a)
MsgBox TypeName(b)
MsgBox TypeName(c)
MsgBox TypeName(d)
MsgBox TypeName(e)

MsgBox TypeName(f)


```



# 4.变量和程序

## 4.1 Option Explicit



```vbscript
option Explicit


dim result

result = MsgBox("Do you want to continue?",vbYesNo)

if result = vbYes  then
	MsgBox "We'll continue..."
else
	MsgBox "Ok,enough..."
end if


```





# 5.流程控制

## 5.1 循环控制



# 6.错误处理和调试



## 6.1 

