> VBS



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



