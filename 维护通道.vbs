rem ===============�����Զ�����================
CurrentPath = createobject("Scripting.FileSystemObject").GetFolder(".").Path
import(CurrentPath&"\myvbsfun.vbs")
rem ==============ʵ������=====================
set myfun=New vbsfun
rem ==============�ж��Ƿ񳬼��û�=============

rem ==============ǰ�ڳ���=====================
call myfun.SyncTime 'ͬ��ʱ��
call myfun.ImportReg(CurrentPath&"\reg.reg")  '�Զ�����ע���
call myfun.RunBat(CurrentPath&"\run.bat")  'ִ��������

rem =============ִ�г���======================
call myfun.Run("c:\windows\notepad.exe",false) 

'===========����ʵ��===========================
set myfun=nothing

'=========================���뺯��=============
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
