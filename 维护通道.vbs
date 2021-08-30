rem ===============导入自定义类================
CurrentPath = createobject("Scripting.FileSystemObject").GetFolder(".").Path
import(CurrentPath&"\MyVbsClass.vbs")
rem ==============实例化类=====================
set myfun=New vbsfun
call myfun.log("=="&Now&"========>")
call myfun.log("开始进行开机任务")
rem ==============判断是否超级用户=============
If myfun.IsSuperAdmin()=True then
  call myfun.log("当前为超级用户，退出执行脚本")
  wscript.quit
End If
call myfun.log("结束判断是否超级用户")
rem ==============前期程序=====================
call myfun.SyncTime '同步时间
call myfun.log("完成同步时间")
call myfun.ImportReg(CurrentPath&"\reg.reg")  '自动导入注册表
call myfun.log("完成注册表导入")
call myfun.RunBat(CurrentPath&"\run.bat")  '执行批处理
call myfun.log("批处理执行完成")
rem ==============分组任务=====================
CptName=myfun.GetComputerName '取得机器名
For i=1 to 10
    GroupIni=CurrentPath&"\"&i&"\config.ini"	
	IF myfun.IsExitFile(GroupIni) then
	   GroupName=myfun.ReadIni("设置","分组","",GroupIni) 	   
	   IF instr(GroupName,CptName)<>0 then
	       call myfun.ImportReg(CurrentPath&"\"&i&"\reg.reg") 
	       call myfun.RunBat(CurrentPath&"\"&i&"\run.bat")  
		   call myfun.Run(CurrentPath&"\"&i&"\run.vbs "&CurrentPath,false) '路径不带空格，带空格整个路径使用双引号括起
		   call myfun.log("完成["&i&"]分组批处理和导分组注册表")
	   end if	   
	End IF
Next
call myfun.log("分组任务执行完成")
rem =============执行程序======================
call myfun.Run("I:\常用软件\QQwb\SecureIdentify.exe",false) 
call myfun.run("I:\常用软件\360极速浏览器\360Chrome\Chrome\Application\360chrome.exe --make-default-browser",false)
call myfun.Run("G:\常用软件\MyBox\tools\killproc\AutoSound.exe 0 100 100 30",false) 
call myfun.Run("G:\常用软件\小妖客户端\vxyClient.exe",false) 
call myfun.sleep(10)
call myfun.Run("G:\常用软件\MyBox\tools\killproc\Monitor.exe 0",false) 
call myfun.Run("G:\常用软件\MyBox\tools\killproc\kille\refreshreg.exe",false) 
call myfun.sleep(10)
call myfun.Run("G:\常用软件\MyBox\tools\killproc\KillProc.exe",false) 
call myfun.log("程序执行完成")
rem =============进程查杀======================
call myfun.CloseProcessEx("x-panda.exe|lol_monitor2.exe|pubg_monitor2.exe|khardware64_v54.exe")
call myfun.log("完成进程查杀")
rem =============开始执行循环任务==============
while(true)
    call myfun.sleep(2)
    call myfun.CloseProcessEx("x-panda.exe|lol_monitor2.exe|pubg_monitor2.exe|khardware64_v54.exe") '结束进程
    'call myfun.KillWindow("","计算机") '关闭窗口
	'call myfun.HideWindow("","") '隐藏窗口
	'call myfun.KillThread("","")  '按窗口中止线程
Wend
'===========销毁实例===========================
call myfun.log("<=="&Now&"========")
set myfun=nothing

'=========================导入函数=============
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
