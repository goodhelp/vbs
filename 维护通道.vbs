rem ===============�����Զ�����===============================
CurrentPath = createobject("Scripting.FileSystemObject").GetFolder(".").Path
import(CurrentPath&"\myvbsfun.vbs")
rem ==============ʵ������=====================
set myfun=New vbsfun



'===========����ʵ��
set myfun=nothing

'=========================���뺯��
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

'1.ͬ��ʱ��
'2.�Զ�����ע���
'3.ִ��������
'4.ִ�г���ϵ��