class vbsfun
	rem ��ʵ����ʱִ�еĴ���
	Public WshShell,FSO,DWX,CurrentPath
	private sub Class_Initialize()
		Set WshShell = WScript.CreateObject("WScript.Shell")
		Set FSO=CreateObject("Scripting.FileSystemObject")
		CurrentPath = createobject("Scripting.FileSystemObject").GetFolder(".").Path
		WshShell.run "regsvr32 /i /s """&CurrentPath&"/dynwrapx.dll""",,true
		Set DWX = CreateObject("DynamicWrapperX")
		'-----windows api--- kernel32.dll----------
		DWX.Register "kernel32 ", "Beep", "i=uu"  
		DWX.Register "kernel32", "GetCommandLine", "r=s"  
		'-----windows api--- user32.dll----------
		DWX.Register "user32", "EnumWindows", "i=ph" 
		DWX.Register "user32", "GetWindowTextW", "i=hpl"
		DWX.Register "user32", "MessageBoxW", "i=hwwu", "r=m"
		DWX.Register "user32", "FindWindow", "i=ss","r=m"
		DWX.Register "user32", "SendMessage", "i=huuu"
		DWX.Register "user32", "ShowWindow", "i=hu", "r=l"
	    DWX.Register "user32", "SetWindowPos", "i=hllllll", "r=l"
		DWX.Register "user32", "PostMessage", "i=hlll", "r=l"
		DWX.Register "user32", "SetWindowText", "i=hs", "r=l"
		DWX.Register "user32", "FindWindowEx", "i=llss", "r=l"
		DWX.Register "user32", "SetCursorPos", "i=ll",  "r=l"
		DWX.Register "user32", "SetWindowRgn","i=hpl","r=l"
		DWX.Register "user32", "GetWindowThreadProcessId","i=hl","r=l"
		DWX.Register "user32", "PostThreadMessage","i=uull","r=l"
		DWX.Register "gdi32", "CreateRectRgn","i=llll","r=p"
		
		
		'https://www.cnblogs.com/jinjiangongzuoshi/p/3905773.html
		'http://dynwrapx.script-coding.com/dwx/pages/dynwrapx.php?lang=en
	end sub

	rem ������ʱִ�еĴ���
	private sub class_terminate()
		WshShell.run "regsvr32 /i /u /s """&CurrentPath&"/dynwrapx.dll""",,true
		Set WshShell=Nothing
		Set FSO=Nothing
		Set DWX=Nothing
	end sub
	
	Rem �����洴��һ����ݷ�ʽ 
	rem ��������ݷ�ʽ����  �����ַ �������в��� ͼ���ַ 
	rem ���� ��
	rem �� call MakeLink("�޼��������","G:\�������\�޼������Ϸ����\Rungame.exe","","G:\�������\�޼������Ϸ����\48731.ico")
	Public Function MakeLink(linkname,linkexe,linkparm,linkico)		
		strDesktop = WshShell.SpecialFolders("Desktop") rem �����ļ��С����桱
		set oShellLink = WshShell.CreateShortcut(strDesktop &"\"& linkname&".lnk")
		oShellLink.TargetPath = linkexe  '��ִ���ļ�·��
		oShellLink.Arguments = linkparm '����Ĳ���
		oShellLink.WindowStyle = 1 '����1Ĭ�ϴ��ڼ������3��󻯼������7��С��
		oShellLink.Hotkey = ""  '��ݼ�
		if IsExitFile(linkico) then
		oShellLink.IconLocation = linkico&", 0" 'ͼ��
		else
		oShellLink.IconLocation = "%SystemRoot%\system32\SHELL32.dll,8"
		end if
		oShellLink.Description = ""  '��ע
		oShellLink.WorkingDirectory = GetExePath(linkexe)  '��ʼλ��
		oShellLink.Save  '���������ݷ�ʽ	
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
		strDesktopFld = objFolderItem.Path
		if link then strDesktopFld=strDesktopFld&"\links"
		Set objURLShortcut = WshShell.CreateShortcut(strDesktopFld & "\"&urlname&".url")
		objURLShortcut.TargetPath = url
		objURLShortcut.Save
		Set objShell=Nothing
	End Function
	
	rem �޸���ҳ
	rem ���� ��ַ
	rem ����
	rem �� call SetHomepage("https://www.baidu.com")
	Public Function SetHomepage(url)
		WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Start Page",url	
	End Function
	
	rem ����exeȡ����·��
	rem ���� ��ȫ·��  
	rem ���� ·��
	rem �� call GetExePath("CProgram FilesInternet Explorer\iexplore.exe")
	Public Function GetExePath(strFileName)
		strFileName=Replace(strFileName,"/","\")
		dim ipos
		ipos=InstrRev(strFileName,"\")
		GetExePath=left(strFileName,ipos)
	End Function

	rem �ж��ļ��Ƿ���� 
	rem ���� �ļ���ַ  
	rem ���� true��false
	rem �� call IsExitFile("c:\abc.txt")
	Public Function IsExitFile(filespec)     
        If FSO.fileExists(filespec) Then         
			IsExitFile=True        
        Else
			IsExitFile=False 
        End If
	End Function 
	
	rem �ж�Ŀ¼�Ƿ���� 
	rem ���� Ŀ¼��ַ �Ƿ񴴽�  
	rem ���� true��false
	rem �� call IsExitDir("c:\abc",true)
	Public Function IsExitDir(DirName,Create)       
        If FSO.folderExists(DirName) Then         
			IsExitDir=True        
        Else
			IsExitDir=False 
			if Create then
				FSO.CreateFolder DirName
			end if
        End If
	End Function
	
	rem �����༶Ŀ¼
	rem ����  ·�� 
	rem ���� ��
	rem ��  call MyCreateFolder("c:\ad\1233\dd")
	Public Sub MyCreateFolder(sPath)
		sPath=Replace(sPath,"/","\")
		if Right(sPath,1)="\" then sPath=left(sPath,len(sPath)-1) 'ɾ��Ŀ¼ĩβ��\
		if(Len(sPath) > 0 And FSO.FolderExists(sPath) = False) Then
			Dim pos, sLeft
			pos = InStrRev(sPath, "\")
			if(pos <> 0) Then
				sLeft = Left(sPath, pos - 1)
				MyCreateFolder sLeft            '�ȴ�����Ŀ¼
			end if
			FSO.CreateFolder sPath              '�ٴ�����Ŀ¼
		end if
	End Sub
	
	rem ����Ŀ¼
	rem ���� ԴĿ¼  Ŀ¼Ŀ¼  �Ƿ��w
	rem ���� �������ļ���
	rem �� call XCopy("c:\123","d:\123",true)
	Public Function XCopy(source, destination, overwrite)
		source=Replace(source,"/","\")
		destination=Replace(destination,"/","\")
		Dim s, d, f, l, CopyCount
		Set s = FSO.GetFolder(source)

		If Not FSO.FolderExists(destination) Then
			FSO.CreateFolder destination
		End If
		Set d = FSO.GetFolder(destination)

		CopyCount = 0
		For Each f In s.Files
			l = d.Path & "\" & f.Name
			If Not FSO.FileExists(l) Or overwrite Then
				If FSO.FileExists(l) Then
					FSO.DeleteFile l, True
				End If
				f.Copy l, True
				CopyCount = CopyCount + 1
			End If
		Next
		For Each f In s.SubFolders
			CopyCount = CopyCount + XCopy(f.Path, d.Path & "\" & f.Name, overwrite)
		Next
		XCopy = CopyCount
	End Function

	rem �����ļ�
	rem ���� Դ�ļ� Ŀ���ļ�  �Ƿ��w
	rem ���� ��
	rem �� call CopyFile("c:\abd\123.txt","d:\323\aaa.txt",true)	
	Public Function CopyFile(sfile,dfile,overwrite)
		if (overwrite and FSO.FileExists(dfile)) then FSO.DeleteFile dfile,true
		if Not FSO.FileExists(GetExePath(dfile)) then
		  MyCreateFolder(GetExePath(dfile))
		end if
		if FSO.fileExists(sFile) then FSO.CopyFile sfile, dfile 
	End Function
	
	rem ɾ���ļ�
	rem ���� Ŀ���ļ�
	rem ���� ��
	rem �� call DelFile("c:\abd\123.txt")	
	Public Function DelFile(sfile)
		if FSO.FileExists(sfile) then FSO.DeleteFile sfile,true
	End Function
	
	rem ɾ��Ŀ¼
	rem ���� Ŀ¼
	rem ���� ��
	rem �� call DelDir("c:\abd\")	
	Public Function DelDir(sPath)
		sPath=Replace(sPath,"/","\")
	    if Right(sPath,1)="\" then sPath=left(sPath,len(sPath)-1)
		if FSO.FolderExists(sPath) then FSO.DeleteFolder sPath
	End Function
	
	rem ���г���
	rem ���� ���� �Ƿ�ȴ�����
	rem ���� ��
	rem �� call Run("c:\abd\123.txt",false)	
	Public Function Run(sPath,wait)
	    if FSO.FileExists(sPath) then
			WshShell.run """"&sPath&"""",,wait
		end if
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
	
	rem ȡ�ñ���IP��ַ
	rem ���� ��
	rem ���ر���IP��ַ
	rem �� call GetIP
	Public Function GetIP
	   ComputerName="."
		Dim objWMIService,colItems,objItem,objAddress
		Set objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\cimv2")
		Set colItems = objWMIService.ExecQuery("Select * From Win32_NetworkAdapterConfiguration Where IPEnabled = True")
		For Each objItem in colItems
			For Each objAddress in objItem.IPAddress
				If objAddress <> "" then
					GetIP = objAddress
					Exit Function
				End If
			Next
		Next
	End Function

	rem ȡ�û�������
	rem ���� ��
	rem ���ر�����������
	rem �� call GetComputerName	
	Public Function GetComputerName
	   ComputerName="."
		Dim objWMIService,colItems,objItem,objAddress
		Set objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\cimv2")
		Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
		For Each objItem in colItems
			GetComputerName = objItem.name
			exit for
		Next	
	End Function
	
	rem ȡ�ò���ϵͳ��
	rem ���� ��
	rem ����  ����ϵͳ��
	rem �� call GetOS	
	Public Function GetOs
	   ComputerName="."
		Dim objWMIService,colItems,objItem,objAddress
		Set objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\cimv2")
		Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
		For Each objItem in colItems
			'GetOs = objItem.Caption&" �汾"& objItem.Version
			if instr(objItem.Version,"6.1")>0 then '6.0��vista 6.1��win7 6.2��win8 10.0��win10
			  GetOS="Win7"
			  exit for
			elseif instr(objItem.Version,"10.0")>0 then
			  GetOs="Win10"
			  exit for
			end if			
		Next	
	End Function
	
	rem ȡ�� ����ϵͳλ��
	rem ���� ��
	rem ����  ����ϵͳλ�� 64λϵͳ����x64 32λϵͳ����x86
	rem �� call X86orX64	
	Public Function X86orX64
	   ComputerName="."
		Dim objWMIService,colItems,objItem,objAddress
		Set objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\cimv2")
		Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem",,48)
		For Each objItem in colItems
		  If InStr(objItem.SystemType, "64") <> 0 Then
		     X86orX64 = "x64" 
		     exit for
		  Else
		     X86orX64 = "x86"
		     exit for
		  End If 		
		Next
	End Function	
	
	rem �ļ�ת��16�����ַ���
	rem ���� �ļ��� 16�����ļ� ��εڶ�������Ϊ�գ�ֱ�ӷ���16�����ַ���
	rem ����16�����ַ��� ���Ϊ�ļ�    16�����ı��ļ���ȿ�ִ�г����һ��
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
			set f =FSO.CreateTextFile(TxtFile,true)
			f.Write node.Text
			f.close
		end if
		Set node = Nothing
		Set xmldom = Nothing
	End Function
	
	rem 16�����ַ���ת�ɿ�ִ���ļ� 
	rem ���� �ַ��� ��ִ���ļ�(��ȫ·��) �Ƿ����ļ� 
	rem ���� ��
	rem �� �ַ������� call BinaryToFile("d:\123.exe","4D5A90000300000004000000FFFF",false)
	rem �� �ı��ļ����� call BinaryToFile("d:\123.exe","d:\123.txt",true)
	Public Sub WriteBinary(exeFile, txtData,IsFile)
		Dim WriteData
		if IsFile then
			Set file = FSO.OpenTextFile(txtData, 1, false)
			WriteData=file.readall
			file.close	
		end if		
		Const adTypeBinary = 1
		Const adSaveCreateOverWrite = 2
		Dim stream, xmldom, node
		Set xmldom = CreateObject("Microsoft.XMLDOM")
		Set node = xmldom.CreateElement("binary")
		node.DataType = "bin.hex"
		node.Text = WriteData
		Set stream = CreateObject("ADODB.Stream")
		stream.Type = adTypeBinary
		stream.Open
		stream.write node.NodeTypedValue
		stream.saveToFile exeFile, adSaveCreateOverWrite
		stream.Close
		Set stream = Nothing
		Set node = Nothing
		Set xmldom = Nothing
	End Sub

	
	rem '��ʱ����	
	rem ����  ��
	rem ���� ��
	rem �� call Sleep(5)
	Public Sub Sleep(sec)
		WScript.sleep sec*1000 
	End sub
	
	rem ����ע����ļ�
	rem ���� �ļ���
	rem ���� ��
	rem �� call ImportReg("d:\1.reg")
	Public Function ImportReg(regFile)
	    if FSO.FileExists(regFile) then
			WshShell.run "regedit.exe /s """&regFile&"""",0
		end if
	End Function	
	
	rem ����bat�ļ�
	rem ���� �ļ���
	rem ���� ��
	rem �� Call RunBat(batFile)
	Public Function RunBat(batFile)
	    if FSO.FileExists(batFile) then
			WshShell.run """"&batFile&"""",0
		end if
	End Function

    rem ����vbs�ļ� 
    rem ���� vbs�ļ�
    rem ���� ��
    rem �� call import("d:\abc.vbs")
    Public Sub import(sFile)
        Dim oFile
        Dim sCode
		if FSO.fileExists(sFile) then 
			Set oFile= FSO.OpenTextFile(sFile, 1)
			With oFile
				sCode= .ReadAll()
				.Close
			End With
			Set oFile= Nothing
		end if
        ExecuteGlobal sCode
    End Sub
	
	rem �ر�ָ������ 
	rem ���� ������
	rem ���� ��
	rem �� call CloseProcess("winrar.exe")
	Public Sub CloseProcess(ExeName)
		WshShell.run "Taskkill /f /im " & ExeName,0
	End Sub

	rem '������  
	rem ���� ������
	rem ���� �����������У�����true
	rem �� Call IsProcess("qq.exe")	
	Public Function IsProcess(ExeName)
		Dim WMI, Obj, Objs,i
		IsProcess = False
		Set WMI = GetObject("WinMgmts:")
		Set Objs = WMI.InstancesOf("Win32_Process")
		For Each Obj In Objs
			If InStr(UCase(ExeName),UCase(Obj.Description)) <> 0 Then
				IsProcess = True
				Exit For
			End If
		Next
		Set Objs = Nothing
		Set WMI = Nothing
	End Function
	
	rem '��������
	rem ���� �����б�����֮����|�ָ�
	rem ���� �����б���ֻҪ��һ�����������з���true
	rem ��	Call IsProcessEx("qq.exe|notepad.exe")
	Public Function IsProcessEx(ExeName)
		Dim WMI, Obj, Objs,ProcessName,i
		IsProcessEx = False
		Set WMI = GetObject("WinMgmts:")
		Set Objs = WMI.InstancesOf("Win32_Process")
		ProcessName=Split(ExeName,"|")
		For Each Obj In Objs
			For i=0 to UBound(ProcessName)
				If InStr(UCase(ProcessName(i)),UCase(Obj.Description)) <> 0 Then
					IsProcessEx = True
					Exit For
				End If
			Next
		Next
		Set Objs = Nothing
		Set WMI = Nothing
	End Function
	
	rem '����������
	rem ���� �����б��м���|�ָ�
	rem ���� ��
	rem ��	call CloseProcessEx("qq.exe��wecat.exe")
	Public Sub CloseProcessEx(ExeName)
		dim ProcessName,CmdCode,i
		ProcessName = Split(ExeName, "|")
		For i=0 to UBound(ProcessName)
			CmdCode=CmdCode & " /im " & ProcessName(i)
			WshShell.run "Taskkill /f" & CmdCode,0
		Next		
	End Sub	
	
	rem ����ƥ��
	
	Public Function RegExpTest(patrn, strng)  
	  Set re = New RegExp  
	  re.Pattern = patrn  
	  re.IgnoreCase = True 
	  re.Global = True 
	  Set Matches = re.Execute(strng)  
	  RegExpTest = Matches.Count  
	  Set re=Nothing
	End Function
	
	rem 'дע���
	rem ���� key ֵ ����
	rem ���� ��
	rem ��	call WriteReg("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\MAIN\Start Page","https://www.baidu.com","")
	Public Sub WriteReg(regkey, value, typeName) 
		If typeName = "" Then
			WshShell.RegWrite regkey, value
		Else
			WshShell.RegWrite regkey, value, typeName
		End If
	End Sub

	rem '��ȡע�������key����������·��
	rem ���� key
	rem ���� ��
	rem ��	call ReadReg("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\MAIN\Start Page")
	Public Function ReadReg(regkey) '
		ReadReg = WshShell.RegRead(regkey)
	End Function

	rem '�ر�ָ�����ⴰ��
	rem ���� ���� ������
	rem ���� ��
	rem ��	call KillWindow("","�ޱ���")
	Public Function KillWindow(classname,winName)
		if len(classname)=0 then classname=0
		if len(winName)=0 then winName=0
		hwnd=DWX.FindWindow(classname,winName)
		DWX.SendMessage hwnd,&H10,0,0 '�رմ���
		'DWX.PostMessage hwnd,&H112,&HF060, 0 '�رմ���
		'DWX.PostMessage hwnd, &H82, 0, 0 '���ٴ���
	   'dim rcSuccess  'ʹ��wscript����alt+F4
	   'rcSuccess = WshShell.AppActivate(winName)
	   'if rcSuccess then WshShell.sendkeys "%{F4}"
	End Function
	
	rem '����ָ�����ⴰ��
	rem ���� ���� ������
	rem ���� ��
	rem ��	call HideWindow("Notepad","")
	Public Function HideWindow(classname,winName)
		if len(classname)=0 then classname=0
		if len(winName)=0 then winName=0
		hwnd=DWX.FindWindow(classname,winName)
	    hrgn =DWX.CreateRectRgn(0,0,0,0)
	    DWX.SetWindowRgn hwnd,hrgn,true '�����ӽ�
		DWX.ShowWindow hwnd,0  '���ش���
	End Function

	rem '���մ�����ֹ�߳�
	rem ���� ���� ������
	rem ���� ��
	rem ��	call KillThread("Notepad","")
	Public Function KillThread(classname,winName)
		if len(classname)=0 then classname=0
		if len(winName)=0 then winName=0
		hwnd=DWX.FindWindow(classname,winName)
	    tid=DWX.GetWindowThreadProcessId(hwnd,0) 'ȡ���߳�ID
		DWX.PostThreadMessage tid,&H12,0,0  '�˳��߳� 
	End Function	

end class