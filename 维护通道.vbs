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

if instr("09,10,11,12,13,14,15,16,17,18",myfun.GetComputerName)<>0 then
 call myfun.RunBat(CurrentPath&"\sub2.bat")  'ִ����������
 call myfun.MakeLink("�޼��������","G:\�������\�޼������Ϸ����\Rungame.exe","","G:\�������\�޼������Ϸ����\48731.ico")
else
 call myfun.RunBat(CurrentPath&"\sub1.bat")  'ִ����������
end if

call myfun.log("������ִ�����")
rem =============ִ�г���======================
call myfun.Run("I:\�������\QQwb\SecureIdentify.exe",false) 
call myfun.run("I:\�������\360���������\360Chrome\Chrome\Application\360chrome.exe --make-default-browser",false)
call myfun.run("G:\�������\MyBox\tools\killproc\AutoSound.exe 0 100 100 30",false)
call myfun.run("G:\�������\MyBox\tools\killproc\UnSee\unsee.exe",false)
call myfun.run("G:\�������\MyBox\tools\killproc\Monitor.exe 0",false)
call myfun.run("G:\�������\MyBox\tools\killproc\kille\refreshreg.exe",false)
call myfun.Sleep(10)
call myfun.run("G:\�������\С���ͻ���\vxyClient.exe",false)
call myfun.Sleep(10)
call myfun.run("G:\�������\MyBox\tools\killproc\KillProc.exe",false)
call myfun.log("����ִ�����")
rem =============���̲�ɱ======================
call myfun.CloseProcessEx("x-panda.exe|lol_monitor2.exe|pubg_monitor2.exe|khardware64_v54.exe")
call myfun.log("��ɽ��̲�ɱ")
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
