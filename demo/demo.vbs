rem �����Զ�����
dim CurrentPath,myfun
CurrentPath = createobject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path
import(createobject("Scripting.FileSystemObject").GetParentFolderName(CurrentPath)&"\lib\MyVbsClass.vbs")
set myfun=New vbsfun
'call myfun.MakeLink("�޼��������","G:\�������\�޼������Ϸ����\Rungame.exe","","G:\�������\�޼������Ϸ����\48731.ico") '���������ݷ�ʽ
'call myfun.MakeUrl("http://www.bnwin.com","������",true) '�ղؼ��������ַ
'call myfun.SetHomepage("https://www.baidu.com") '����ie��ҳ
'msgbox myfun.GetExePath("C:\Program Files\Internet Explorer\iexplore.exe") '����·��ȡĿ¼
'msgbox myfun.IsExitFile("c:\abc.txt") '�ж��ļ��Ƿ����
'msgbox myfun.IsExitDir("c:\abc",true) '�ж�Ŀ¼�Ƿ����
'call myfun.MyCreateFolder("c:\abc\1233\dd")  '����Ŀ¼�ɶ༶
'call myfun.XCopy("D:\dump","d:\456",true) '��Ŀ¼ �༶
'call myfun.CopyFile("C:\Windows\win.ini","d:\323\aaa.txt",true)  '�����ļ�
'call myfun.DelFile("c:\abd\123.txt")  'ɾ���ļ�
'call myfun.DelDir("c:\abd\")  'ɾ��Ŀ¼
'Call myfun.Run("""C:\Program Files (x86)\Internet Explorer\iexplore.exe""  http://www.bnwin.com",false)  '·�����пո�Ҫ�����Űѿո�·�����𣬻����ļ��Ƿ���ڣ���������ִ��dos����
'msgbox myfun.ping("192.168.0.1")  'ping�Ƿ�����
'MsgBox myfun.GetMac   'ȡ������mac��ַ
'MsgBox myfun.GetIP   'ȡ�ñ���ip��ַ
'MsgBox myfun.GetComputerName   'ȡ�û�����
'MsgBox myfun.GetOS  'ȡ�ò���ϵͳ��win7����win10
'MsgBox myfun.X86orX64  'ϵͳ��64λ����32λ
'call myfun.ReadBinary("c:\windows\notepad.exe","d:\123.txt") '���ļ�����16�����ַ����ı��ļ�
'call myfun.WriteBinary("d:\123.exe","d:\123.txt",true)  '��16�����ı��ļ���ԭΪ��ִ���ļ�
'call myfun.Sleep(5)  '��ʱ5��
'call myfun.ImportReg(".\reg.reg")  '����ע���
'call myfun.RunBat(".\test.bat")  '�����������ļ�
'myfun.Runcmd "dir c:\ >c:\1.txt" '����dos����
'call myfun.import("d:\abc.vbs")  '����vbs�ļ�
'call myfun.CloseProcess("SunloginRemote.exe")  '�ر�ָ������
'call myfun.IsProcess("qq.exe")	 '���ָ�������Ƿ����
'call myfun.IsProcessEx("qq.exe|notepad.exe")  'ָ��ָ���б�����Ƿ����
'call myfun.CloseProcessEx("qq.exe|wechat.exe")  '���������б�
'call myfun.RegExpTest("sdf|456","123456789")  '�����ж��Ƿ����ָ������
'call myfun.KillWindow("","�����")  '�ر�ָ�����������Ĵ���
'call myfun.WriteReg("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\MAIN\Start Page","https://www.baidu.com","")  'дע���ֵ
'MsgBox myfun.ReadReg("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\MAIN\Start Page") '��ע���ֵ
'call myfun.HideWindow("Notepad","") '��������������ش��� δ�ر�
'call myfun.KillThread("Notepad","")  '�������������ֹ��ʾ�˴��ڵ��߳� ����ǵ��̣߳���������˳� �����ڹرչ�洰��
'call myfun.DownFile("http://gh.api.99988866.xyz/github.com/Chuyu-Team/Dism-Multi-language/releases/download/v10.1.1001.10/Dism++10.1.1001.10_d4ba4eb035254b3326d6adc6638bc9c8daea7018.zip","d:\dism.zip")  '����Զ���ļ� 
'call myfun.SyncTime  'ͬ��ʱ��
'MsgBox myfun.GetSysRunTime() 'ȡ��ϵͳ����ʱ�� �ػ�ʱ�� ����ʱ��ϳ���������ѧϰ��ѯϵͳ��־
'call myfun.log("����")   'д��־�ļ� ��־�ļ��ڵ�ǰ�ű��£�����������
'MsgBox myfun.IsSuperAdmin() '�ж������Ƿ��г����û�
'MsgBox myfun.FromUnixTime("1630211522",8) 'unixʱ��ת����ʱ��
'msgbox myfun.ToUnixTime(now, 8)  '�ѵ�ǰʱ��ת��unixʱ��
'myfun.WriteIni "�ڵ�","����","ֵ","d:\123.ini"   'дINI�ļ�
'msgbox myfun.ReadIni("�ڵ�","����","Ĭ��ֵ","d:\123.ini") '��ini�ļ�
'call myfun.CreatLink("d:\pg","C:\Program Files (x86)")  ''��C�̳���Ŀ¼ӳ�䵽D��pgĿ¼,����D:\pg�൱�ڷ���C:\Program Files (x86)����
'msgbox "ϵͳ�����ˣ�"&myfun.GetOsRunTime&"����"
'call myfun.SysVolme  '��ϵͳ�����������
'call myfun.Depolicy("��ֹ���ö˿�")  '����������ȫ���Խ�ֹ135 137 139 445 3389�˿�
'call myfun.CloseMonitor(2) '2�ر���ʾ -1����ʾ��
'call myfun.RunAu3("E:\�������\vbs\demo\monitor.au3",false) '�ָ���ʾ����������
'call myfun.RestMonitor() '�ָ���ʾ����������

dim cptname
cptname=myfun.GetComputerName

'MsgBox(cptname&"���н���")
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