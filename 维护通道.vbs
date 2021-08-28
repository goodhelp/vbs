rem ===============导入自定义类================
CurrentPath = createobject("Scripting.FileSystemObject").GetFolder(".").Path
import(CurrentPath&"\myvbsfun.vbs")
rem ==============实例化类=====================
set myfun=New vbsfun
rem ==============判断是否超级用户=============

rem ==============前期程序=====================
call myfun.SyncTime '同步时间
call myfun.ImportReg(CurrentPath&"\reg.reg")  '自动导入注册表
call myfun.RunBat(CurrentPath&"\run.bat")  '执行批处理

rem =============执行程序======================
call myfun.Run("c:\windows\notepad.exe",false) 

'===========销毁实例===========================
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
