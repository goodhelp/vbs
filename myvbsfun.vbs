set myfun=New vbsfun
'call myfun.CopyFile("d:\Users\Administrator\Desktop\myvbsfun.vbs","d:\aa\123.vbs",true)
'call myfun.Run("regedit",true)
MsgBox("����")

class vbsfun
	rem ��ʵ����ʱִ�еĴ���
	private sub Class_Initialize()

	end sub

	rem ������ʱִ�еĴ���
	private sub class_terminate()

	end sub

	Rem �����洴��һ�����±���ݷ�ʽ 
	rem ��������ݷ�ʽ����  �����ַ �������в��� ͼ���ַ 
	rem ���� ��
	rem �� call MakeLink("�޼��������.lnk","G:\�������\�޼������Ϸ����\Rungame.exe","","G:\�������\�޼������Ϸ����\48731.ico")
	Public Function MakeLink(linkname,linkexe,linkparm,linkico)
		Set WshShell = WScript.CreateObject("WScript.Shell")
		strDesktop = WshShell.SpecialFolders("Desktop") rem �����ļ��С����桱
		set oShellLink = WshShell.CreateShortcut(strDesktop &"\"& linkname)
		oShellLink.TargetPath = linkexe  '��ִ���ļ�·��
		oShellLink.Arguments = linkparm '����Ĳ���
		oShellLink.WindowStyle = 1 '����1Ĭ�ϴ��ڼ������3��󻯼������7��С��
		oShellLink.Hotkey = ""  '��ݼ�
		if IsExitAFile(linkico) then
		oShellLink.IconLocation = linkico&", 0" 'ͼ��
		else
		oShellLink.IconLocation = "%SystemRoot%\system32\SHELL32.dll,8"
		end if
		oShellLink.Description = ""  '��ע
		oShellLink.WorkingDirectory = GetExePath(linkexe)  '��ʼλ��
		oShellLink.Save  '���������ݷ�ʽ	
		Set WshShell=Nothing
		Set oShellLink=Nothing
	End Function
	
	rem ����exeȡ����·��
	rem ���� ��ȫ·��  
	rem ���� ·��
	rem �� call GetExePat("CProgram FilesInternet Explorer\iexplore.exe")
	Public Function GetExePath(strFileName)
		strFileName=Replace(strFileName,"/","\")
		dim ipos
		ipos=InstrRev(strFileName,"\")
		GetExePath=left(strFileName,ipos)
	End Function

	rem �ж��ļ��Ƿ���� 
	rem ���� �ļ���ַ  
	rem ���� true��false
	rem �� call IsExitAFile("c:\abc.txt")
	Public Function IsExitAFile(filespec)
        Dim fso
        Set fso=CreateObject("Scripting.FileSystemObject")        
        If fso.fileExists(filespec) Then         
			IsExitAFile=True        
        Else
			IsExitAFile=False 
        End If
		Set fso=Nothing
	End Function 
	
	rem �ж�Ŀ¼�Ƿ���� 
	rem ���� Ŀ¼��ַ �Ƿ񴴽�  
	rem ���� true��false
	rem �� call IsExitDir("c:\abc",true)
	Public Function IsExitDir(DirName,Create)
        Dim fso
        Set fso=CreateObject("Scripting.FileSystemObject")        
        If fso.folderExists(DirName) Then         
			IsExitDir=True        
        Else
			IsExitDir=False 
			if Create then
				fso.CreateFolder DirName
			end if
        End If
		Set fso=Nothing
	End Function
	
	rem �����༶Ŀ¼
	rem ����  ·�� 
	rem ���� ��
	rem ��  call MyCreateFolder("c:\ad\1233\dd")
	Public Sub MyCreateFolder(sPath)
		sPath=Replace(sPath,"/","\")
		if Right(sPath,1)="\" then sPath=left(sPath,len(sPath)-1) 'ɾ��Ŀ¼ĩβ��\
		Dim fso
		set fso = CreateObject("Scripting.FileSystemObject")
		if(Len(sPath) > 0 And fso.FolderExists(sPath) = False) Then
			Dim pos, sLeft
			pos = InStrRev(sPath, "\")
			if(pos <> 0) Then
				sLeft = Left(sPath, pos - 1)
				MyCreateFolder sLeft            '�ȴ�����Ŀ¼
			end if
			fso.CreateFolder sPath              '�ٴ�����Ŀ¼
		end if
		set fso = Nothing
	End Sub
	
	rem ����Ŀ¼
	rem ���� ԴĿ¼  Ŀ¼Ŀ¼  �Ƿ��w
	rem ���� �������ļ���
	rem �� call XCopy("c:\123" "d:\123",true)
	Public Function XCopy( source, destination, overwrite)
		source=Replace(source,"/","\")
		destination=Replace(destination,"/","\")
		Dim fso,s, d, f, l, CopyCount
		set fso = CreateObject("Scripting.FileSystemObject")
		Set s = fso.GetFolder(source)

		If Not fso.FolderExists(destination) Then
			fso.CreateFolder destination
		End If
		Set d = fso.GetFolder(destination)

		CopyCount = 0
		For Each f In s.Files
			l = d.Path & "\" & f.Name
			If Not fso.FileExists(l) Or overwrite Then
				If fso.FileExists(l) Then
					fso.DeleteFile l, True
				End If
				f.Copy l, True
				CopyCount = CopyCount + 1
			End If
		Next
		For Each f In s.SubFolders
			CopyCount = CopyCount + XCopy(f.Path, d.Path & "\" & f.Name, overwrite)
		Next
		XCopy = CopyCount
		Set fso=Nothing
	End Function

	rem �����ļ�
	rem ���� Դ�ļ� Ŀ���ļ�  �Ƿ��w
	rem ���� ��
	rem �� call CopyFile("c:\abd\123.txt","d:\323\aaa.txt",true)	
	Public Function CopyFile(sfile,dfile,overwrite)
		dim fso
		set fso = CreateObject("Scripting.FileSystemObject")
		if (overwrite and fso.FileExists(dfile)) then fso.DeleteFile dfile,true
		if Not fso.FileExists(GetExePath(dfile)) then
		  MyCreateFolder(GetExePath(dfile))
		end if
		fso.CopyFile sfile, dfile 
		set fso = nothing
	End Function
	
	rem ɾ���ļ�
	rem ���� Ŀ���ļ�
	rem ���� ��
	rem �� call DelFile("c:\abd\123.txt")	
	Public Function DelFile(sfile)
		dim fso
		set fso = CreateObject("Scripting.FileSystemObject")
		if fso.FileExists(sfile) then fso.DeleteFile sfile,true
		set fso = nothing
	End Function
	
	rem ɾ��Ŀ¼
	rem ���� Ŀ¼
	rem ���� ��
	rem �� call DelDir("c:\abd\123.txt")	
	Public Function DelDir(sPath)
		sPath=Replace(sPath,"/","\")
	    if Right(sPath,1)="\" then sPath=left(sPath,len(sPath)-1)
		dim fso
		set fso = CreateObject("Scripting.FileSystemObject")
		if fso.FolderExists(sPath) then fso.DeleteFolder sPath
		set fso = nothing
	End Function
	
	rem ���г���
	rem ���� ���� �Ƿ�ȴ�����
	rem ���� ��
	rem �� call Run("c:\abd\123.txt",false)	
	Public Function Run(sPath,wait)
		dim shell
		Set shell = Wscript.createobject("WScript.shell")
		shell.run """"&sPath&"""",,wait
		set shell = nothing
	End Function
	
	rem ping�����Ƿ�����
	rem ���� IP��ַ 
	rem ����true��false
	rem �� call ping("192.168.0.1")
	Public Function Ping(strComputer)
		Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
		Set colItems = objWMIService.ExecQuery _
			("Select * from Win32_PingStatus " & _
				"Where Address = '" & strComputer & "'")
		For Each objItem in colItems
			If objItem.StatusCode = 0 Then 
				Ping=true 
			else
				Ping=false  			
			End If
		Next
		Set objWMIService=Nothing
	End Function
	
end class