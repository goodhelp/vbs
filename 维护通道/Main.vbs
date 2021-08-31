rem ===============导入自定义类================
dim vbsPath
vbsPath = createobject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path'脚本当前目录
import(createobject("Scripting.FileSystemObject").GetParentFolderName(vbsPath)&"\lib\MyVbsClass.vbs")
rem ==============实例化类=====================
dim myfun
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
call myfun.ImportReg(vbsPath&"\reg.reg")  '自动导入注册表
call myfun.log("完成注册表导入")
call myfun.RunBat(vbsPath&"\run.bat")  '执行批处理
call myfun.log("批处理执行完成")
rem ==============分组任务=====================
dim CptName,g,GroupIni,GroupName,Plist
CptName=myfun.GetComputerName '取得机器名
For g=1 to 10
    GroupIni=vbsPath&"\"&g&"\config.ini"	
	IF myfun.IsExitFile(GroupIni) then
	   GroupName=myfun.ReadIni("设置","分组","",GroupIni) 	   
	   IF instr(GroupName,CptName)<>0 then
	       call myfun.ImportReg(vbsPath&"\"&g&"\reg.reg") 
	       call myfun.RunBat(vbsPath&"\"&g&"\run.bat")  
		   if myfun.IsExitFile(vbsPath&"\"&g&"\run.vbs") then
		      import(vbsPath&"\"&g&"\run.vbs")
		   end if
		   call myfun.log("完成["&g&"]分组批处理和导分组注册表")
	   end if	   
	End IF
Next
call myfun.log("分组任务执行完成")
rem =============执行程序======================
if myfun.IsExitFile(vbsPath&"\Runpg.vbs") then
    import(vbsPath&"\Runpg.vbs")
end if
rem =============进程查杀======================
Plist="x-panda.exe|lol_monitor2.exe|pubg_monitor2.exe|khardware64_v54.exe"
call myfun.CloseProcessEx(Plist)
call myfun.log("完成进程查杀")
rem =============开始执行循环任务==============
if myfun.IsExitFile(vbsPath&"\taskloop.vbs") then
    import(vbsPath&"\taskloop.vbs")
end if
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
