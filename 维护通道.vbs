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

rem =============执行程序======================
call myfun.Run("c:\windows\notepad.exe",false) 

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
