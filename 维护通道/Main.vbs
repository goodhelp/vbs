rem ===============�����Զ�����================
dim vbsPath
vbsPath = createobject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path'�ű���ǰĿ¼
import(createobject("Scripting.FileSystemObject").GetParentFolderName(vbsPath)&"\lib\MyVbsClass.vbs")
rem ==============ʵ������=====================
dim myfun
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
call myfun.ImportReg(vbsPath&"\reg.reg")  '�Զ�����ע���
call myfun.log("���ע�����")
call myfun.RunBat(vbsPath&"\run.bat")  'ִ��������
call myfun.log("������ִ�����")
rem ==============��������=====================
dim CptName,g,GroupIni,GroupName,Plist
CptName=myfun.GetComputerName 'ȡ�û�����
For g=1 to 10
    GroupIni=vbsPath&"\"&g&"\config.ini"	
	IF myfun.IsExitFile(GroupIni) then
	   GroupName=myfun.ReadIni("����","����","",GroupIni) 	   
	   IF instr(GroupName,CptName)<>0 then
	       call myfun.ImportReg(vbsPath&"\"&g&"\reg.reg") 
	       call myfun.RunBat(vbsPath&"\"&g&"\run.bat")  
		   if myfun.IsExitFile(vbsPath&"\"&g&"\run.vbs") then
		      import(vbsPath&"\"&g&"\run.vbs")
		   end if
		   call myfun.log("���["&g&"]����������͵�����ע���")
	   end if	   
	End IF
Next
call myfun.log("��������ִ�����")
rem =============ִ�г���======================
if myfun.IsExitFile(vbsPath&"\Runpg.vbs") then
    import(vbsPath&"\Runpg.vbs")
end if
rem =============���̲�ɱ======================
Plist="x-panda.exe|lol_monitor2.exe|pubg_monitor2.exe|khardware64_v54.exe"
call myfun.CloseProcessEx(Plist)
call myfun.log("��ɽ��̲�ɱ")
rem =============��ʼִ��ѭ������==============
if myfun.IsExitFile(vbsPath&"\taskloop.vbs") then
    import(vbsPath&"\taskloop.vbs")
end if
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
