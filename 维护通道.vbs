rem ===============导入自定义类===============================
CurrentPath = createobject("Scripting.FileSystemObject").GetFolder(".").Path
import(CurrentPath&"\myvbsfun.vbs")
rem ==============实例化类=====================
set myfun=New vbsfun



'===========销毁实例
set myfun=nothing

'=========================导入函数
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

'1.同步时间
'2.自动导入注册表
'3.执行批处理
'4.执行程序系列