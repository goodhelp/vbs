class vbsfun
	' ��ʵ����ʱִ�еĴ���
	Public WSH,FSO,DWX,AU3,CurrentPath
	private sub Class_Initialize()
		Set WSH = WScript.CreateObject("WScript.Shell")
		Set FSO=CreateObject("Scripting.FileSystemObject")
		Set DIC = CreateObject("Scripting.Dictionary")
		CurrentPath = createobject("Scripting.FileSystemObject").GetFile(Wscript.ScriptFullName).ParentFolder.Path
		WSH.run "regsvr32 /i /s """&createobject("Scripting.FileSystemObject").GetParentFolderName(CurrentPath)&"\lib\dynwrapx.dll""",,true
		WSH.run "regsvr32 /i /s """&createobject("Scripting.FileSystemObject").GetParentFolderName(CurrentPath)&"\lib\AutoItX3.dll""",,true
		Set DWX = CreateObject("DynamicWrapperX")
		Set AU3 = WScript.CreateObject("AutoItX3.Control")
		'-----windows api--- kernel32.dll---------- 
		'http://dynwrapx.script-coding.com/dwx/pages/dynwrapx.php?lang=en
		'https://omen999.developpez.com/tutoriels/vbs/dynawrapperx-v2-1/
		'https://blog.csdn.net/yxp_xa/article/details/73320759
		'https://www.jb51.net/shouce/vbs/vtoriVBScript.htm 'vbs�̳�
		'http://www.bathome.net/thread-4068-1-2.html 'wmic�̳�
		DWX.Register "kernel32 ", "Beep", "i=uu"  
		DWX.Register "kernel32", "GetCommandLine", "r=s" 
        DWX.Register "kernel32", "GetPrivateProfileString","i=sssSus", "r=u" 
		DWX.Register "kernel32", "WritePrivateProfileString","i=ssss", "r=l" 
		DWX.Register "kernel32", "GetTickCount","r=l"
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
		DWX.Register "user32", "GetWindowThreadProcessId","i=hL","r=l"
		DWX.Register "user32", "PostThreadMessage","i=uull","r=l"
		DWX.Register "user32", "SendMessageTimeout","i=hlhhlll"
		DWX.Register "user32", "MonitorFromWindow","i=hl","r=h"
		DWX.Register "user32", "GetDesktopWindow","r=h"
		'--------------gdi32.dll-----------------------------
		DWX.Register "gdi32", "CreateRectRgn","i=llll","r=p"	
		DWX.Register "Dxva2","RestoreMonitorFactoryDefaults","i=h"
		DWX.Register "Dxva2","GetNumberOfPhysicalMonitorsFromHMONITOR","i=hp"
		
	end sub

	' ������ʱִ�еĴ���
	private sub class_terminate()
		WSH.run "regsvr32 /i /u /s """&createobject("Scripting.FileSystemObject").GetParentFolderName(CurrentPath)&"\lib\dynwrapx.dll""",,true
		WSH.run "regsvr32 /i /u /s """&createobject("Scripting.FileSystemObject").GetParentFolderName(CurrentPath)&"\lib\AutoItX3.dll""",,true
		Set WSH=Nothing
		Set FSO=Nothing
		Set DIC=Nothing
		Set DWX=Nothing
		Set AU3=Nothing
	end sub
	
	' �����洴��һ����ݷ�ʽ 
	' ��������ݷ�ʽ����  �����ַ �������в��� ͼ���ַ 
	' ���� ��
	' �� call MakeLink("�޼��������","G:\�������\�޼������Ϸ����\Rungame.exe","","G:\�������\�޼������Ϸ����\48731.ico")
	Public Function MakeLink(linkname,linkexe,linkparm,linkico)	
		dim strDesktop,oShellLink
		strDesktop = WSH.SpecialFolders("Desktop") rem �����ļ��С����桱
		set oShellLink = WSH.CreateShortcut(strDesktop &"\"& linkname&".lnk")
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
	
	' �ղؼ������ַ
	' ����:��ַ ������� �Ƿ񴴽����ղؼ���
	' ���� ��
	' �� call MakeUrl("http://www.bnwin.com","������",true)	
	Public Function MakeUrl(url,urlname,link)
		Const ADMINISTRATIVE_TOOLS = 6
		dim objShell,objFolder,objFolderItem,strDesktopFld,objURLShortcut
		Set objShell = CreateObject("Shell.Application")
		Set objFolder = objShell.Namespace(ADMINISTRATIVE_TOOLS)
		Set objFolderItem = objFolder.Self 		
		strDesktopFld = objFolderItem.Path
		if link then strDesktopFld=strDesktopFld&"\links"
		Set objURLShortcut = WSH.CreateShortcut(strDesktopFld & "\"&urlname&".url")
		objURLShortcut.TargetPath = url
		objURLShortcut.Save
		Set objShell=Nothing
	End Function
	
	' �޸���ҳ
	' ���� ��ַ
	' ����
	' �� call SetHomepage("https://www.baidu.com")
	Public Function SetHomepage(url)
		WriteReg "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Start Page",url,""	
		WriteReg "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main\Start Page",url,""
		WSH.Run "cmd.exe /c gpupdate /force",0,false 
		WSH.Run "RunDll32.exe USER32.DLL,UpdatePerUserSystemParameters",0,false 
		DWX.SendMessageTimeout &HFFFF,&H1A,0,0,0,1000,0
	End Function
	
	' ����exeȡ����·��
	' ���� ��ȫ·��  
	' ���� ·��
	' �� call GetExePath("CProgram FilesInternet Explorer\iexplore.exe")
	Public Function GetExePath(strFileName)
		strFileName=Replace(strFileName,"/","\")
		dim ipos
		ipos=InstrRev(strFileName,"\")
		GetExePath=left(strFileName,ipos)
	End Function

	' �ж��ļ��Ƿ���� 
	' ���� �ļ���ַ  
	' ���� true��false
	' �� call IsExitFile("c:\abc.txt")
	Public Function IsExitFile(filespec)     
        If FSO.fileExists(filespec) Then         
			IsExitFile=True        
        Else
			IsExitFile=False 
        End If
	End Function 
	
	' �ж�Ŀ¼�Ƿ���� 
	' ���� Ŀ¼��ַ �Ƿ񴴽�  
	' ���� true��false
	' �� call IsExitDir("c:\abc",true)
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
	
	' �����༶Ŀ¼
	' ����  ·�� 
	' ���� ��
	' ��  call MyCreateFolder("c:\ad\1233\dd")
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
	
	' ����Ŀ¼
	' ���� ԴĿ¼  Ŀ¼Ŀ¼  �Ƿ��w
	' ���� �������ļ���
	' �� call XCopy("c:\123","d:\123",true)
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

	' �����ļ�
	' ���� Դ�ļ� Ŀ���ļ�  �Ƿ��w
	' ���� ��
	' �� call CopyFile("c:\abd\123.txt","d:\323\aaa.txt",true)	
	Public Function CopyFile(sfile,dfile,overwrite)
		if (overwrite and FSO.FileExists(dfile)) then FSO.DeleteFile dfile,true
		if Not FSO.FileExists(GetExePath(dfile)) then
		  MyCreateFolder(GetExePath(dfile))
		end if
		if FSO.fileExists(sFile) then FSO.CopyFile sfile, dfile 
	End Function
	
	' ɾ���ļ�
	' ���� Ŀ���ļ�
	' ���� ��
	' �� call DelFile("c:\abd\123.txt")	
	Public Function DelFile(sfile)
		if FSO.FileExists(sfile) then FSO.DeleteFile sfile,true
	End Function
	
	' ɾ��Ŀ¼
	' ���� Ŀ¼
	' ���� ��
	' �� call DelDir("c:\abd\")	
	Public Function DelDir(sPath)
		sPath=Replace(sPath,"/","\")
	    if Right(sPath,1)="\" then sPath=left(sPath,len(sPath)-1)
		if FSO.FolderExists(sPath) then FSO.DeleteFolder sPath
	End Function
	
	' ���г��� ·�����ո���Ҫ��˫��������·��
	' ���� ���� �Ƿ�ȴ�����
	' ���� ��
	' �� call Run("c:\abd\123.txt",false)	
	Public Function Run(sPath,wait)
	    on error resume next
		err.clear
	    dim ExeName,IsRun,Exepath,i
	    ExeName = Split(sPath, " ")	'�ָ����ո��·��
		For i=0 to UBound(ExeName)
		   if i=0 then
		     Exepath=ExeName(i)
		   else
		     Exepath=Exepath&" "&ExeName(i) '������ϳɴ��ո��·��
		   end if
		   Exepath=Replace(Exepath,"""","")  
           if FSO.FileExists(Exepath) then
		     IsRun=True
			 exit for
		   end if
        next
        if IsRun then		
		   WSH.run sPath,,wait
		end if
		if err.number<>0 then
		  log("ִ��Run����"&Err.Source&Err.Description&Err.Number)
		end if
	End Function
	
	' ping�����Ƿ�����
	' ���� IP��ַ 
	' ����true��false
	' �� call ping("192.168.0.1")
	Public Function Ping(strComputer)
		dim objWMIService,colItems
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
	
	' ȡ������MAC��ַ
	' ���� ��
	' ���ر���mac��ַ
	' �� call GetMac	
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
	
	' ȡ�ñ���IP��ַ
	' ���� ��
	' ���ر���IP��ַ
	' �� call GetIP
	Public Function GetIP
	   dim ComputerName
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

	' ȡ�û�������
	' ���� ��
	' ���ر�����������
	' �� call GetComputerName	
	Public Function GetComputerName
	   dim ComputerName
	   ComputerName="."
		Dim objWMIService,colItems,objItem,objAddress
		Set objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\cimv2")
		Set colItems = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
		For Each objItem in colItems
			GetComputerName = objItem.name
			exit for
		Next	
	End Function
	
	' ȡ�ò���ϵͳ��
	' ���� ��
	' ����  ����ϵͳ��
	' �� call GetOS	
	Public Function GetOs
	   dim ComputerName
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
	
	' ȡ�� ����ϵͳλ��
	' ���� ��
	' ����  ����ϵͳλ�� 64λϵͳ����x64 32λϵͳ����x86
	' �� call X86orX64	
	Public Function X86orX64
	   dim ComputerName
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
	
	' �ļ�ת��16�����ַ���
	' ���� �ļ��� 16�����ļ� ��εڶ�������Ϊ�գ�ֱ�ӷ���16�����ַ���
	' ����16�����ַ��� ���Ϊ�ļ�    16�����ı��ļ���ȿ�ִ�г����һ��
	' �������ַ��� call ReadBinary("c:\windows\notepad.exe","")
	' �������ı��ļ� call ReadBinary("c:\windows\notepad.exe","d:\123.txt")
	Public Function ReadBinary(FileName,TxtFile)
		Const adTypeBinary = 1
		Dim stream, xmldom, node,f
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
	
	' 16�����ַ���ת�ɿ�ִ���ļ� 
	' ���� �ַ��� ��ִ���ļ�(��ȫ·��) �Ƿ����ļ� 
	' ���� ��
	' �� �ַ������� call BinaryToFile("d:\123.exe","4D5A90000300000004000000FFFF",false)
	' �� �ı��ļ����� call BinaryToFile("d:\123.exe","d:\123.txt",true)
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

	' ����Զ���ļ�������
	' ���� Զ�̵�ַ �����ļ�
	' ���� ��
	' �� call DownFile("https://dl.360safe.com/360sd/360sd_x64_std_5.0.0.8183C.exe","d:\360sd.exe")
	Public Function DownFile(UrlFile,SaveFile)
	    dim xPost,sGet
		Set xPost=CreateObject("Microsoft.XMLHTTP")
		xPost.Open "get",UrlFile,0
		xPost.Send()
		Set sGet=CreateObject("ADODB.Stream")
		sGet.type=1
		sGet.Mode=3
		sGet.Open()
		sGet.Write(xPost.responseBody)
		sGet.SaveToFile SaveFile,2
	    Set sGet=Nothing
		Set xPost=Nothing
	End Function
	
	' '��ʱ����	
	' ����  ��
	' ���� ��
	' �� call Sleep(5)
	Public Sub Sleep(sec)
		WScript.sleep sec*1000 
	End sub
	
	' ����ע����ļ�
	' ���� �ļ���
	' ���� ��
	' �� call ImportReg("d:\1.reg")
	Public Function ImportReg(regFile)
	    if FSO.FileExists(regFile) then
			WSH.run "regedit.exe /s """&regFile&"""",0
		end if
	End Function	
	
	' ����bat�ļ�
	' ���� �ļ���
	' ���� ��
	' �� Call RunBat(batFile)
	Public Function RunBat(batFile)
	    if FSO.FileExists(batFile) then
			WSH.run """"&batFile&"""",0
		end if
	End Function
	
	' ����dos����
	' ���� dos����
	' ���� ��
	' �� Call RunCmd(batstr)
	Public Function RunCmd(batstr)
		WSH.run "cmd.exe /c "&batstr,0
	End Function

    ' ����vbs�ļ� 
    ' ���� vbs�ļ�
    ' ���� ��
    ' �� call import("d:\abc.vbs")
    Public Sub import(sFile)
        Dim oFile,sCode
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
	
	' �ر�ָ������ 
	' ���� ������
	' ���� ��
	' �� call CloseProcess("winrar.exe")
	Public Sub CloseProcess(ExeName)
	    if IsProcess(ExeName) then
		  WSH.run "Taskkill /f /im " & ExeName,0
		end if
	End Sub

	' '������  
	' ���� ������
	' ���� �����������У�����true
	' �� Call IsProcess("qq.exe")	
	Public Function IsProcess(ExeName)
		Dim WMI, Obj, Objs
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
	
	' '��������
	' ���� �����б�����֮����|�ָ�
	' ���� �����б���ֻҪ��һ�����������з���true
	' ��	Call IsProcessEx("qq.exe|notepad.exe")
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
	
	' '����������
	' ���� �����б��м���|�ָ�
	' ���� ��
	' ��	call CloseProcessEx("qq.exe��wecat.exe")
	Public Sub CloseProcessEx(ExeName)
		dim ProcessName,CmdCode,i
		ProcessName = Split(ExeName, "|")
		For i=0 to UBound(ProcessName)
		    if IsProcess(ProcessName(i)) then  '������̴���
			  CmdCode=CmdCode & " /im " & ProcessName(i)			  
			end if
		Next
        IF len(CmdCode)>0 then
           WSH.run "Taskkill /f" & CmdCode,0
        End If		   
	End Sub	
	
	' ����ƥ��
	
	Public Function RegExpTest(patrn, strng) 
      dim re,Matches	
	  Set re = New RegExp  
	  re.Pattern = patrn  
	  re.IgnoreCase = True 
	  re.Global = True 
	  Set Matches = re.Execute(strng)  
	  RegExpTest = Matches.Count  
	  Set re=Nothing
	End Function
	
	' 'дע���
	' ���� key ֵ ����
	' ���� ��
	' ��	call WriteReg("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\MAIN\Start Page","https://www.baidu.com","")
	Public Sub WriteReg(regkey, value, typeName) 
	    on error resume next
		err.clear
		If typeName = "" Then
			WSH.RegWrite regkey, value
		Else
			WSH.RegWrite regkey, value, typeName
		End If
		if err.number<>0 then
		  log("дע������"&Err.Source&Err.Description&Err.Number)
		end if		
	End Sub

	' '��ȡע�������key����������·��
	' ���� key
	' ���� ��
	' ��	call ReadReg("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\MAIN\Start Page")
	Public Function ReadReg(regkey) '
	    on error resume next
		err.clear
		ReadReg = WSH.RegRead(regkey)
		if err.number<>0 then
		  ReadReg=false
		 end if 
		'if err.number<>0 then
		'  log("��ȡע������"&Err.Source&Err.Description&Err.Number)
		'end if
	End Function

	' '�ر�ָ�����ⴰ��
	' ���� ���� ������
	' ���� ��
	' ��	call KillWindow("","�ޱ���")
	Public Function KillWindow(classname,winName)
	    dim hwnd
		if len(classname)=0 then classname=0
		if len(winName)=0 then winName=0
		hwnd=DWX.FindWindow(classname,winName)
		if hwnd<>0 then
		  DWX.SendMessage hwnd,&H10,0,0 '�رմ���
		'DWX.PostMessage hwnd,&H112,&HF060, 0 '�رմ���
		'DWX.PostMessage hwnd, &H82, 0, 0 '���ٴ���
		end if
	   'dim rcSuccess  'ʹ��wscript����alt+F4
	   'rcSuccess = WSH.AppActivate(winName)
	   'if rcSuccess then WSH.sendkeys "%{F4}"
	End Function
	
	' '����ָ�����ⴰ��
	' ���� ���� ������
	' ���� ��
	' ��	call HideWindow("Notepad","")
	Public Function HideWindow(classname,winName)
	    dim hwnd,hrgn
		if len(classname)=0 then classname=0
		if len(winName)=0 then winName=0
		hwnd=DWX.FindWindow(classname,winName)
		if hwnd<>0 then
	      hrgn =DWX.CreateRectRgn(0,0,0,0)
	      DWX.SetWindowRgn hwnd,hrgn,true '�����ӽ�
		  DWX.ShowWindow hwnd,0  '���ش���
		End If
	End Function

	' '���մ�����ֹ�߳�
	' ���� ���� ������
	' ���� ��
	' ��	call KillThread("Notepad","")
	Public Function KillThread(classname,winName)
	    dim hwnd,tid,dwProcessID
		if len(classname)=0 then classname=0
		if len(winName)=0 then winName=0
		hwnd=DWX.FindWindow(classname,winName)
		if hwnd<>0 then
			tid=DWX.GetWindowThreadProcessId(hwnd,dwProcessID) 'ȡ���߳�TID ����IDΪdwProcessID 
			DWX.PostThreadMessage tid,&H12,0,0  '�˳��߳� 
		end if
	End Function	
	
	' ͬ������ʱ��
	' ���� ��
	' ���� 
	' �� call SyncTime
	Public Sub SyncTime()
        On error resume next
		err.clear
        dim url,xmlHTTP,objRegEx,Contents,colMatches,strMatch,strMatches,dtmNewDate,strMatches1,dtmNewTime	
	    url = "http://free.timeanddate.com/clock/i1jyoa52/n236/tt0/tw0/tm3/td2/th1/tb4" 
		'url = "http://api.m.taobao.com/rest/api3.do?api=mtop.common.getTimestamp" '���Ա�ʱ��api
		'Instantiate
		Set xmlHTTP = CreateObject("MSXML2.ServerXMLHTTP") 
		Set objRegEx = CreateObject("VBScript.RegExp")
		XMLhttp.setTimeouts 5000,5000,5000,15000
		'Make Request
		xmlHTTP.open "GET", url, false 
		xmlHTTP.send ""
		'Wait for Response
		xmlHTTP.waitForResponse()
		objRegEx.Global = True 
		'If status is 200, then it's OK
		If xmlHTTP.status = 200 then
		   Contents=xmlHTTP.responseText
		'get date info
		   objRegEx.Pattern = "\d{2,2}/\d{2,2}/\d{4,4}"
			Set colMatches = objRegEx.Execute(Contents) 
			If colMatches.Count > 0 Then
			 For Each strMatch in colMatches 
				 strMatches = strMatches & strMatch.Value
			 Next
			End If
			if len(strMatches)=0 then exit Sub
			dtmNewDate = FormatDateTime(strMatches,D)
		'set date on local computer
		   WSH.Run "cmd.exe /c date " & dtmNewDate,0 
		'get time info
		   objRegEx.Pattern = "\d{2,2}:\d{2,2}:\d{2,2}"
			Set colMatches = objRegEx.Execute(Contents) 
			If colMatches.Count > 0 Then
			 For Each strMatch in colMatches 
				 strMatches1 = strMatches1 & strMatch.Value
			 Next
			End If
			dtmNewTime = strMatches1
			'wscript.echo dtmNewTime
		'set time on local computer
		   WSH.Run "cmd.exe /c time " & dtmNewTime,0 
		End if
		if err.number<>0 then
		  log("��ȡ����ʱ�����"&Err.Source&Err.Description&Err.Number)
		end if
	End Sub
	
	' ȡ��ϵͳ����ʱ�� �ػ�ʱ�� ����ʱ��ϳ���������ѧϰ��ѯϵͳ��־
	' ���� ��
	' ���� 
	' �� call GetSysRunTime
	Public Function GetSysRunTime()
	    dim strComputer,objWMIService,colLoggedEvents,objEvent,Flag
		strComputer = "."
		Set objWMIService = GetObject("winmgmts:" _
		& "{impersonationLevel=impersonate}!\\" _
		& strComputer & "\root\cimv2")
		Set colLoggedEvents = objWMIService.ExecQuery _
		("Select * from Win32_NTLogEvent " _
		& "Where Logfile = 'System' And EventCode = '6005' Or EventCode = '6006'")
		For Each objEvent In colLoggedEvents
		  Flag = Flag + 1
		  If Flag = 1 Then
			GetSysRunTime= "���ο���ʱ��: " & FormatWMIUTC(objEvent.TimeWritten)
		  ElseIf Flag = 2 Then
			GetSysRunTime=GetSysRunTime& " �ϴιػ�ʱ��: " & FormatWMIUTC(objEvent.TimeWritten)
		  ElseIf Flag = 3 Then
			GetSysRunTime=GetSysRunTime& " �ϴο���ʱ��: " & FormatWMIUTC(objEvent.TimeWritten)
			Exit For
		  End If
		Next
    End Function
	
	'����ϵͳ���е�ʱ�� ����
	'���� 
	'���� ����
	'�� GetOsRunTime()
	Public Function GetOsRunTime()
	   dim r
	   r=DWX.GetTickCount
	   GetOsRunTime=int(r/1000/60)
	End Function
	
	' ��ʽ��wmiʱ��
	'FormatUTC
	Public Function FormatWMIUTC(WMIDateString)
	  dim DS,i
	  DS = " // :: "
	  FormatWMIUTC = Left(WMIDateString,2)
	  For i = 2 To 7
		FormatWMIUTC = FormatWMIUTC & Mid(WMIDateString, i * 2 - 1, 2) & Mid(DS,i,1)
	  Next
	  'FormatWMIUTC = Mid(WMIDateString, 1, 4) & "��" _
	  '      & Mid(WMIDateString, 5, 2) & "��" _
	  '      & Mid(WMIDateString, 7, 2) & "�� " _
	  '      & Mid (WMIDateString, 9, 2) & ":" _
	  '      & Mid(WMIDateString, 11, 2) & ":" _
	  '      & Mid(WMIDateString,13, 2)
	End Function
	
	'�ѱ�׼ʱ��ת��ΪUNIXʱ���
	'������strTime:Ҫת����ʱ�䣻intTimeZone����ʱ���Ӧ��ʱ��       
	'����ֵ��strTime�����1970��1��1����ҹ0�㾭��������       
	'ʾ����ToUnixTime("2008-5-23 10:51:0", +8)������ֵΪ1211511060       
	Public Function ToUnixTime(strTime, intTimeZone)       
		If IsEmpty(strTime) or Not IsDate(strTime) Then strTime = Now       
		If IsEmpty(intTimeZone) or Not isNumeric(intTimeZone) Then intTimeZone = 0       
		ToUnixTime = DateAdd("h",-intTimeZone,strTime)       
		ToUnixTime = DateDiff("s","1970-1-1 0:0:0", ToUnixTime)       
	End Function      
		  
	'��UNIXʱ���ת��Ϊ��׼ʱ��       
	'������intTime:Ҫת����UNIXʱ�����intTimeZone����ʱ�����Ӧ��ʱ��       
	'����ֵ��intTime������ı�׼ʱ��       
	'ʾ����FromUnixTime("1211511060", +8)������ֵ2008-5-23 10:51:0       
	Public Function FromUnixTime(intTime, intTimeZone)       
		If IsEmpty(intTime) Or Not IsNumeric(intTime) Then      
			FromUnixTime = Now()       
		   Exit Function      
		End If      
		If IsEmpty(intTime) Or Not IsNumeric(intTimeZone) Then intTimeZone = 0       
		FromUnixTime = DateAdd("s", intTime, "1970-1-1 0:0:0")       
		FromUnixTime = DateAdd("h", intTimeZone, FromUnixTime)       
	End Function	
	
	' ������־
	' ���� ��־���ݣ�
	' ����ֵ  ��
	' ʾ��: log("�¼���־")
	Public Function log(logstr)	
        dim logfile	,ObjLog
	    logfile=CurrentPath&"\log-"&year(Now)&"-"&Month(Now)&"-"&day(Now)&".txt"
	    if fso.FileExists(logfile) then
           Set ObjLog = FSO.OpenTextFile(LogFile,8)		   
		else 
          Set ObjLog = FSO.CreateTextFile(logfile)
		end if
		ObjLog.Write vbCrLf&"["&Now&"]  ��־���ݣ�"&logstr
		ObjLog.close
		set ObjLog=Nothing
	End Function
	
	'�жϵ�ǰ�Ƿ����̳����û�
	'��������
	'����ֵ true ��false
	'ʾ����IsSuperAdmin
	Public Function IsSuperAdmin()
		'[��ά��ʦ/�̻���ʦ/����Win����]
		dim AdminValue
		AdminValue=ReadReg("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\iCafe8\Admin")
		If AdminValue=1 then
		    IsSuperAdmin=true
			exit Function
		End IF
		'[������]
		AdminValue=ReadReg("HKEY_LOCAL_MACHINE\SOFTWARE\EYOOCLIENTSTATUS\SuperLogin")
		If AdminValue=1 then
		    IsSuperAdmin=true
			exit Function
		End IF
		'[�Ƹ���]
		AdminValue=ReadReg("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\SuperAdmin")
		If AdminValue=1 then
		    IsSuperAdmin=true
			exit Function
		End IF
		'[������]
		AdminValue=ReadReg("HKEY_LOCAL_MACHINE\SYSTEM\nVos\SupperMode")
		If AdminValue=1 then
		    IsSuperAdmin=true
			exit Function
		End IF
		'[����]
		AdminValue=ReadReg("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Richdisk\ClientFlag")
		If AdminValue=1 then
		    IsSuperAdmin=true
			exit Function
		End IF
        IsSuperAdmin=false
	End Function
	
	'���ܣ���ȡINI�ļ�
	'������ �ڵ�,������Ĭ��ֵ��ini�ļ�
	'����ֵ:��ֵ
	'ʾ����ReadIni("�ڵ�","����","Ĭ��ֵ","d:\123.ini")
	Public Function ReadIni(iSection,iKey,dValue,iFile)
	  if len(iSection)<>0 and len(iFile)<>0 then 
		  dim vStr
		  vStr=Space(255)
		  Call DWX.GetPrivateProfileString(iSection,iKey,dValue,vStr,255,iFile)
		  ReadIni=vStr
		  Set vStr=Nothing
      End If		  
	End Function
	
	'���ܣ�дINI�ļ�
	'�������ڵ㣬��������ֵ��ini�ļ�
	'����ֵ:д��ɹ�����1 ���򷵻�0
	'ʾ����WriteIni "�ڵ�","����","ֵ","d:\123.ini"
	Public Function WriteIni(iSection,iKey,iValue,iFile)
	  if len(iSection)<>0 and len(iKey)<>0 and len(iFile)<>0 then 
	     WriteIni=DWX.WritePrivateProfileString(iSection,iKey,iValue,iFile)
	  End if
	End Function
	
	'���ܣ�����Ŀ¼����
	'��������Ŀ¼ ԴĿ¼
	'����ֵ:��
	'ʾ��  CreatLink("C:\Program Files\Adobe\Photoshop","D:\Program Files\Adobe\Photoshop") '��D��photoshopӳ�䵽C��,ʵ�ʳ����ļ������D��
	Public Function CreatLink(NewDir,OldDir)
	  if FSO.FolderExists(NewDir) then 
	     WSH.run "cmd.exe /c rd "& NewDir,,false
	  end if
	  if FSO.FolderExists(OldDir) then
	     WSH.run "cmd.exe /c mklink /d """& NewDir & """ """  & OldDir&"""" ,,false
	  End if
	End Function
	
	'���� ����ϵͳ����Ϊ���
	'���� �ޱ���
	'����ֵ 
	'ʾ����SysVolme ԭ�� ������������������� AF��Ȼ����һ������URL������빤�߶� %97%AF ���н��룬�õ����ַ��� ��������
	'��Ĭ�����������ҳ Sendkeys "��"  �򿪡��ҵĵ��ԡ� Sendkeys "��"  �򿪡���������  Sendkeys "��" 
	Public Sub SysVolme()
	  dim i
	  for i=0 to 50
	     WSH.Sendkeys "��"  '������ Sendkeys "��" ���� Sendkeys "��"
	  next	
	End Sub
	
	'���� ��������ȫ���Խ�ֹ135 137 139 445 3389�˿�
	'���� ��������
	'����ֵ ��
	'ʾ�� call myfun.Depolicy("��ֹ���ö˿�")  '����������ȫ���Խ�ֹ135 137 139 445 3389�˿�
	Public Sub Depolicy(plname)
	    WSH.Run "netsh ipsec static del all",0,true
	    WSH.Run "netsh  ipsec static add policy name="&plname,0,true
		WSH.Run "netsh  ipsec static add filterlist name=denyip",0,true
		WSH.Run "netsh  ipsec static add filter filterlist=denyip srcaddr=Any dstaddr=Me dstport=135 protocol=TCP",0,true
		WSH.Run "netsh  ipsec static add filter filterlist=denyip srcaddr=Any dstaddr=Me dstport=135 protocol=UDP",0,true
		WSH.Run "netsh  psec static add filter filterlist=denyip srcaddr=Any dstaddr=Me dstport=137 protocol=TCP",0,true
		WSH.Run "netsh  ipsec static add filter filterlist=denyip srcaddr=Any dstaddr=Me dstport=137 protocol=UDP",0,true
		WSH.Run "netsh  ipsec static add filter filterlist=denyip srcaddr=Any dstaddr=Me dstport=139 protocol=TCP",0,true
		WSH.Run "netsh  ipsec static add filter filterlist=denyip srcaddr=Any dstaddr=Me dstport=139 protocol=UDP",0,true
		WSH.Run "netsh  ipsec static add filter filterlist=denyip srcaddr=Any dstaddr=Me dstport=445 protocol=TCP",0,true
		WSH.Run "netsh  ipsec static add filter filterlist=denyip srcaddr=Any dstaddr=Me dstport=445 protocol=UDP",0,true
		WSH.Run "netsh  ipsec static add filter filterlist=denyip srcaddr=Any dstaddr=Me dstport=3389 protocol=TCP",0,true
		WSH.Run "netsh  ipsec static add filter filterlist=denyip srcaddr=Any dstaddr=Me dstport=3389 protocol=UDP",0,true
		'WSH.Run "netsh ipsec static add filter filterlist=denyip srcaddr=Me dstaddr=Any dstport=3505 protocol=TCP",0,true  '��ֹ�������ӵ��κε�ַ��3505�˿�
		'WSH.Run "netsh ipsec static add filter filterlist=denyip srcaddr=Me dstaddr=192.168.0.236",0,true '��ֹ�������ӵ�192.168.0.236
		WSH.Run "netsh  ipsec static add filteraction name=denyact action=block",0,true
		WSH.Run "netsh  ipsec static add rule name=killport policy="&plname&" filterlist=denyip filteraction=denyact",0,true
		WSH.Run "netsh  ipsec static set policy name="&plname&" assign=y",0,true	
	End Sub
	
	'���� �رջ����ʾ��
	'���� -1�� ��2�ر�
	'����ֵ ��
	'ʾ��  CloseMonitor(2)
	Public Sub CloseMonitor(turn)
	   DWX.PostMessage &HFFFF,&H112,&HF170&,turn
	End Sub
	
	'�ж�com���Ƿ�װ
	Public Function ComExist(ComName)
		On Error Resume Next
		Set CreateTest = CreateObject(ComName)
		ComExist = CBool(Err.Number = 0)
		On Error Goto 0
	End Function	
	
	'���� ����autoit3�ű��ļ�
	'���� �ű��ļ�ȫ·��
	'����ֵ
	'ʾ��  RunAu3("E:\�������\vbs\demo\monitor.au3") '�ָ���ʾ����������
    Public Sub RunAu3(au3File)
	  if FSO.fileExists(au3File) then
	     WSH.run """"&createobject("Scripting.FileSystemObject").GetParentFolderName(CurrentPath)&"\lib\AutoIt3.exe"" "&au3File,0,false
      End if	  
	End Sub	
	
end class