rem �����Զ�����
import("E:\�������\vbs\myvbsfun.vbs")
set myfun=New vbsfun
'call myfun.MakeLink("�޼��������","G:\�������\�޼������Ϸ����\Rungame.exe","","G:\�������\�޼������Ϸ����\48731.ico") '���������ݷ�ʽ
'call myfun.MakeUrl("http://www.bnwin.com","������",true) '�ղؼ��������ַ
'call myfun.SetHomepage("http://www.bnwin.com") '����ie��ҳ
'call myfun.GetExePath("C:\Program Files\Internet Explorer\iexplore.exe") '����·��ȡĿ¼
'call myfun.IsExitFile("c:\abc.txt") '�ж��ļ��Ƿ����
'call myfun.IsExitDir("c:\abc",true) '�ж�Ŀ¼�Ƿ����
'call myfun.MyCreateFolder("c:\abc\1233\dd")  '����Ŀ¼�ɶ༶
'call myfun.XCopy("D:\dump","d:\456",true) '��Ŀ¼ �༶
'call myfun.CopyFile("C:\Windows\win.ini","d:\323\aaa.txt",true)  '�����ļ�
'call myfun.DelFile("c:\abd\123.txt")  'ɾ���ļ�
'call myfun.DelDir("c:\abd\")  'ɾ��Ŀ¼
'call myfun.Run("c:\windows\notepad.exe",false)	 '���г���
'call myfun.ping("192.168.0.1")  'ping�Ƿ�����
'call myfun.GetMac   'ȡ������mac��ַ
'call myfun.GetIP   'ȡ�ñ���ip��ַ
'call myfun.GetComputerName   'ȡ�û�����
'call myfun.GetOS  'ȡ�ò���ϵͳ��win7����win10
'call myfun.X86orX64  'ϵͳ��64λ����32λ
'call myfun.ReadBinary("c:\windows\notepad.exe","d:\123.txt") '���ļ�����16�����ַ����ı��ļ�
'call myfun.WriteBinary("d:\123.exe","d:\123.txt",true)  '��16�����ı��ļ���ԭΪ��ִ���ļ�
'call myfun.Sleep(5)  '��ʱ5��
'call myfun.ImportReg(".\reg.reg")  '����ע���
'call myfun.RunBat(".\test.bat")  '�����������ļ�
'call myfun.import("d:\abc.vbs")  '����vbs�ļ�
'call myfun.CloseProcess("SunloginRemote.exe")  '�ر�ָ������
'call myfun.IsProcess("qq.exe")	 '���ָ�������Ƿ����
'call myfun.IsProcessEx("qq.exe|notepad.exe")  'ָ��ָ���б�����Ƿ����
'call myfun.CloseProcessEx("qq.exe��wecat.exe")  '���������б�
'call myfun.RegExpTest("sdf|456","123456789")  '�����ж��Ƿ����ָ������

dim cptname
cptname=myfun.GetComputerName

MsgBox(cptname&"���н���")
set myfun=nothing

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