rem ===============�����Զ�����================
CurrentPath = createobject("Scripting.FileSystemObject").GetFolder(".").Path
import(CurrentPath&"\MyVbsClass.vbs")
rem ==============ʵ������=====================
set myfun=New vbsfun
call myfun.log("=="&Now&"========>")
call myfun.log("��ʼ���п�������")
rem ==============�ж��Ƿ񳬼��û�=============
If myfun.IsSuperAdmin()=True then
  call myfun.log("��ǰΪ�����û����˳�ִ�нű�")
  wscript.quit
End If
call myfun.log("�����ж��Ƿ񳬼��û�")
rem ==============ǰ�ڳ���=====================
call myfun.SyncTime 'ͬ��ʱ��
call myfun.log("���ͬ��ʱ��")
call myfun.ImportReg(CurrentPath&"\reg.reg")  '�Զ�����ע���
call myfun.log("���ע�����")
call myfun.RunBat(CurrentPath&"\run.bat")  'ִ��������
call myfun.log("������ִ�����")

rem =============ִ�г���======================
call myfun.Run("c:\windows\notepad.exe",false) 

'===========����ʵ��===========================
call myfun.log("<=="&Now&"========")
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
