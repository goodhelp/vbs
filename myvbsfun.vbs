set myfun=New vbsfun
'call myfun.CopyFile("d:\Users\Administrator\Desktop\myvbsfun.vbs","d:\aa\123.vbs",true)
'call myfun.Run("regedit",true)
'call myfun.SetHomepage("http://www.bnwin.com")
'call myfun.ReadBinary("c:\windows\notepad.exe","d:\123.txt")
call myfun.BinaryToFile("d:\123.txt","d:\123.exe",true)
MsgBox(myfun.GetMac)


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
	rem �� call MakeLink("�޼��������","G:\�������\�޼������Ϸ����\Rungame.exe","","G:\�������\�޼������Ϸ����\48731.ico")
	Public Function MakeLink(linkname,linkexe,linkparm,linkico)
		Set WshShell = WScript.CreateObject("WScript.Shell")
		strDesktop = WshShell.SpecialFolders("Desktop") rem �����ļ��С����桱
		set oShellLink = WshShell.CreateShortcut(strDesktop &"\"& linkname&".lnk")
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
	
	rem �ղؼ������ַ
	rem ����:��ַ ������� �Ƿ񴴽����ղؼ���
	rem ���� ��
	rem �� call MakeUrl("http://www.bnwin.com","������",true)	
	Public Function MakeUrl(url,urlname,link)
		Const ADMINISTRATIVE_TOOLS = 6
		Set objShell = CreateObject("Shell.Application")
		Set objFolder = objShell.Namespace(ADMINISTRATIVE_TOOLS)
		Set objFolderItem = objFolder.Self 		
		Set objShell = WScript.CreateObject("WScript.Shell")
		strDesktopFld = objFolderItem.Path
		if link then strDesktopFld=strDesktopFld&"\links"
		Set objURLShortcut = objShell.CreateShortcut(strDesktopFld & "\"&urlname&".url")
		objURLShortcut.TargetPath = url
		objURLShortcut.Save
		Set objShell=Nothing
	End Function
	
	rem �޸���ҳ
	rem ���� ��ַ
	rem ����
	rem �� SetHomepage("https://www.baidu.com")
	Public Function SetHomepage(url)
		dim oShell
		Set oShell = CreateObject("WScript.Shell")
		oShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Start Page",url	
		set oShell=Nothing
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
	
	rem ȡ������MAC��ַ
	rem ���� ��
	rem ���ر���mac��ַ
	rem �� call GetMac	
	Public Function GetMac
		Dim mc,mo
		Set mc=GetObject("Winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")
		For Each mo In mc
		If mo.IPEnabled=True Then
		  GetMac=mo.MacAddress
		Exit For
		End If
		Next
	End Function
	
	rem �ļ�ת��16�����ַ��� ���� https://blog.csdn.net/yuman198629/article/details/8595694
	rem ���� �ļ��� 16�����ļ� ��εڶ�������Ϊ�գ�ֱ�ӷ���16�����ַ���
	rem ����16�����ַ���
	rem �������ַ��� call ReadBinary("c:\windows\notepad.exe","")
	rem �������ı��ļ� call ReadBinary("c:\windows\notepad.exe","d:\123.txt")
	Public Function ReadBinary(FileName,TxtFile)
		Const adTypeBinary = 1
		Dim stream, xmldom, node
		Set xmldom = CreateObject("Microsoft.XMLDOM")
		Set node = xmldom.CreateElement("binary")
		node.DataType = "bin.hex"
		Set stream = CreateObject("ADODB.Stream")
		stream.Type = adTypeBinary
		stream.Open
		stream.LoadFromFile FileName
		node.NodeTypedValue = stream.Read
		stream.Close
		Set stream = Nothing
		if len(TxtFile)=0 then
			ReadBinary = node.Text
		else
			Set FSO = CreateObject("Scripting.FileSystemObject")
			set f =fso.CreateTextFile(TxtFile,true)
			f.Write node.Text
			f.close
			set FSO=Nothing
		end if
		Set node = Nothing
		Set xmldom = Nothing
	End Function
	
	rem 16�����ַ���ת�ɿ�ִ���ļ� 
	rem ���� �ַ��� ��ִ���ļ�(��ȫ·��) �Ƿ����ļ� 
	rem ���� ��
	rem �� �ַ������� call BinaryToFile("4D5A90000300000004000000FFFF","d:\123.exe",false)
	rem �� �ı��ļ����� call BinaryToFile("d:\123.txt","d:\123.exe",true)
	Public Function BinaryToFile(WriteData,dropFileName,isfile)
		Set FSO = CreateObject("Scripting.FileSystemObject")
	    if isfile then
			Set file = fso.OpenTextFile(WriteData, 1, false)
			WriteData=file.readall
			file.close
		end if
		If FSO.FileExists(dropFileName)=False Then
		Set FileObj = FSO.CreateTextFile(dropFileName, True)
		For i = 1 To Len(WriteData) Step 2
		   FileObj.Write Chr(CLng("&H" & Mid(WriteData,i,2)))
		Next
		FileObj.Close
		End If
		Set FSO=Nothing
	End Function

end class