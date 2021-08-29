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

if instr("09,10,11,12,13,14,15,16,17,18",myfun.GetComputerName)<>0 then
 call myfun.RunBat(CurrentPath&"\sub2.bat")  '执行子批处理
 call myfun.MakeLink("罗技鼠标设置","G:\常用软件\罗技鼠标游戏驱动\Rungame.exe","","G:\常用软件\罗技鼠标游戏驱动\48731.ico")
else
 call myfun.RunBat(CurrentPath&"\sub1.bat")  '执行子批处理
end if

call myfun.log("批处理执行完成")
rem =============执行程序======================
call myfun.Run("I:\常用软件\QQwb\SecureIdentify.exe",false) 
call myfun.run("I:\常用软件\360极速浏览器\360Chrome\Chrome\Application\360chrome.exe --make-default-browser",false)
call myfun.run("G:\常用软件\MyBox\tools\killproc\AutoSound.exe 0 100 100 30",false)
call myfun.run("G:\常用软件\MyBox\tools\killproc\UnSee\unsee.exe",false)
call myfun.run("G:\常用软件\MyBox\tools\killproc\Monitor.exe 0",false)
call myfun.run("G:\常用软件\MyBox\tools\killproc\kille\refreshreg.exe",false)
call myfun.Sleep(10)
call myfun.run("G:\常用软件\小妖客户端\vxyClient.exe",false)
call myfun.Sleep(10)
call myfun.run("G:\常用软件\MyBox\tools\killproc\KillProc.exe",false)
call myfun.log("程序执行完成")
rem =============进程查杀======================
call myfun.CloseProcessEx("x-panda.exe|lol_monitor2.exe|pubg_monitor2.exe|khardware64_v54.exe")
call myfun.log("完成进程查杀")
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
