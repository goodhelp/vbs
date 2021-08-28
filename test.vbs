rem 导入自定义类
import("E:\软件工程\vbs\myvbsfun.vbs")
set myfun=New vbsfun
'call myfun.MakeLink("罗技鼠标设置","G:\常用软件\罗技鼠标游戏驱动\Rungame.exe","","G:\常用软件\罗技鼠标游戏驱动\48731.ico") '创建桌面快捷方式
'call myfun.MakeUrl("http://www.bnwin.com","百脑问",true) '收藏夹栏添加网址
'call myfun.SetHomepage("http://www.bnwin.com") '设置ie主页
'call myfun.GetExePath("C:\Program Files\Internet Explorer\iexplore.exe") '根据路径取目录
'call myfun.IsExitFile("c:\abc.txt") '判断文件是否存在
'call myfun.IsExitDir("c:\abc",true) '判断目录是否存在
'call myfun.MyCreateFolder("c:\abc\1233\dd")  '创建目录可多级
'call myfun.XCopy("D:\dump","d:\456",true) '拷目录 多级
'call myfun.CopyFile("C:\Windows\win.ini","d:\323\aaa.txt",true)  '拷贝文件
'call myfun.DelFile("c:\abd\123.txt")  '删除文件
'call myfun.DelDir("c:\abd\")  '删除目录
'call myfun.Run("c:\windows\notepad.exe",false)	 '运行程序
'call myfun.ping("192.168.0.1")  'ping是否在线
'call myfun.GetMac   '取得网卡mac地址
'call myfun.GetIP   '取得本机ip地址
'call myfun.GetComputerName   '取得机器名
'call myfun.GetOS  '取得操作系统是win7还是win10
'call myfun.X86orX64  '系统是64位还是32位
'call myfun.ReadBinary("c:\windows\notepad.exe","d:\123.txt") '把文件生成16进制字符串文本文件
'call myfun.WriteBinary("d:\123.exe","d:\123.txt",true)  '把16进制文本文件还原为可执行文件
'call myfun.Sleep(5)  '延时5秒
'call myfun.ImportReg(".\reg.reg")  '导入注册表
'call myfun.RunBat(".\test.bat")  '运行批处理文件
'call myfun.import("d:\abc.vbs")  '导入vbs文件
'call myfun.CloseProcess("SunloginRemote.exe")  '关闭指定进程
'call myfun.IsProcess("qq.exe")	 '检查指定进程是否存在
'call myfun.IsProcessEx("qq.exe|notepad.exe")  '指定指定列表进程是否存在
'call myfun.CloseProcessEx("qq.exe｜wecat.exe")  '结束进程列表
'call myfun.RegExpTest("sdf|456","123456789")  '正则判断是否存在指定内容

dim cptname
cptname=myfun.GetComputerName

MsgBox(cptname&"运行结束")
set myfun=nothing

Sub import(sFile)
	Dim oFSO, oFile,sCode
	Set oFSO	= CreateObject("Scripting.FileSystemObject")
	Set oFile	= oFSO.OpenTextFile(sFile, 1)
	With oFile
		sCode	= .ReadAll()
		.Close
	End With
	Set oFile	= Nothing
	Set oFSO	= Nothing
	ExecuteGlobal sCode
End Sub