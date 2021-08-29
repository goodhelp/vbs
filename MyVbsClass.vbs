class vbsfun
	' 类实例化时执行的代码
	Public WshShell,FSO,DWX,CurrentPath
	private sub Class_Initialize()
		Set WshShell = WScript.CreateObject("WScript.Shell")
		Set FSO=CreateObject("Scripting.FileSystemObject")
		Set Dict = CreateObject("Scripting.Dictionary")
		CurrentPath = createobject("Scripting.FileSystemObject").GetFolder(".").Path
		WshShell.run "regsvr32 /i /s """&CurrentPath&"/dynwrapx.dll""",,true
		Set DWX = CreateObject("DynamicWrapperX")
		'-----windows api--- kernel32.dll---------- 
		'http://dynwrapx.script-coding.com/dwx/pages/dynwrapx.php?lang=en
		'https://omen999.developpez.com/tutoriels/vbs/dynawrapperx-v2-1/
		'https://www.jb51.net/shouce/vbs/vtoriVBScript.htm
		DWX.Register "kernel32 ", "Beep", "i=uu"  
		DWX.Register "kernel32", "GetCommandLine", "r=s" 
        DWX.Register "kernel32", "GetPrivateProfileString","i=sssSus", "r=u" 
		DWX.Register "kernel32", "WritePrivateProfileString","i=ssss", "r=l" 
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
		'--------------gdi32.dll-----------------------------
		DWX.Register "gdi32", "CreateRectRgn","i=llll","r=p"	
		
	end sub

	' 类销毁时执行的代码
	private sub class_terminate()
		WshShell.run "regsvr32 /i /u /s """&CurrentPath&"/dynwrapx.dll""",,true
		Set WshShell=Nothing
		Set FSO=Nothing
		Set Dict=Nothing
		Set DWX=Nothing
	end sub
	
	' 在桌面创建一个快捷方式 
	' 参数：快捷方式名称  程序地址 程序运行参数 图标地址 
	' 返回 无
	' 例 call MakeLink("罗技鼠标设置","G:\常用软件\罗技鼠标游戏驱动\Rungame.exe","","G:\常用软件\罗技鼠标游戏驱动\48731.ico")
	Public Function MakeLink(linkname,linkexe,linkparm,linkico)		
		strDesktop = WshShell.SpecialFolders("Desktop") rem 特殊文件夹“桌面”
		set oShellLink = WshShell.CreateShortcut(strDesktop &"\"& linkname&".lnk")
		oShellLink.TargetPath = linkexe  '可执行文件路径
		oShellLink.Arguments = linkparm '程序的参数
		oShellLink.WindowStyle = 1 '参数1默认窗口激活，参数3最大化激活，参数7最小化
		oShellLink.Hotkey = ""  '快捷键
		if IsExitFile(linkico) then
		oShellLink.IconLocation = linkico&", 0" '图标
		else
		oShellLink.IconLocation = "%SystemRoot%\system32\SHELL32.dll,8"
		end if
		oShellLink.Description = ""  '备注
		oShellLink.WorkingDirectory = GetExePath(linkexe)  '起始位置
		oShellLink.Save  '创建保存快捷方式	
		Set oShellLink=Nothing
	End Function
	
	' 收藏夹添加网址
	' 参数:网址 快捷名称 是否创建在收藏夹栏
	' 返回 无
	' 例 call MakeUrl("http://www.bnwin.com","百脑问",true)	
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
	
	' 修改主页
	' 参数 网址
	' 返回
	' 例 call SetHomepage("https://www.baidu.com")
	Public Function SetHomepage(url)
		WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Start Page",url	
	End Function
	
	' 根据exe取所在路径
	' 参数 完全路径  
	' 返回 路径
	' 例 call GetExePath("CProgram FilesInternet Explorer\iexplore.exe")
	Public Function GetExePath(strFileName)
		strFileName=Replace(strFileName,"/","\")
		dim ipos
		ipos=InstrRev(strFileName,"\")
		GetExePath=left(strFileName,ipos)
	End Function

	' 判断文件是否存在 
	' 参数 文件地址  
	' 返回 true或false
	' 例 call IsExitFile("c:\abc.txt")
	Public Function IsExitFile(filespec)     
        If FSO.fileExists(filespec) Then         
			IsExitFile=True        
        Else
			IsExitFile=False 
        End If
	End Function 
	
	' 判断目录是否存在 
	' 参数 目录地址 是否创建  
	' 返回 true或false
	' 例 call IsExitDir("c:\abc",true)
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
	
	' 创建多级目录
	' 参数  路径 
	' 返回 无
	' 例  call MyCreateFolder("c:\ad\1233\dd")
	Public Sub MyCreateFolder(sPath)
		sPath=Replace(sPath,"/","\")
		if Right(sPath,1)="\" then sPath=left(sPath,len(sPath)-1) '删除目录末尾的\
		if(Len(sPath) > 0 And FSO.FolderExists(sPath) = False) Then
			Dim pos, sLeft
			pos = InStrRev(sPath, "\")
			if(pos <> 0) Then
				sLeft = Left(sPath, pos - 1)
				MyCreateFolder sLeft            '先创建父目录
			end if
			FSO.CreateFolder sPath              '再创建本目录
		end if
	End Sub
	
	' 拷贝目录
	' 参数 源目录  目录目录  是否覆蓋
	' 返回 拷贝的文件数
	' 例 call XCopy("c:\123","d:\123",true)
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

	' 复制文件
	' 参数 源文件 目标文件  是否覆蓋
	' 返回 无
	' 例 call CopyFile("c:\abd\123.txt","d:\323\aaa.txt",true)	
	Public Function CopyFile(sfile,dfile,overwrite)
		if (overwrite and FSO.FileExists(dfile)) then FSO.DeleteFile dfile,true
		if Not FSO.FileExists(GetExePath(dfile)) then
		  MyCreateFolder(GetExePath(dfile))
		end if
		if FSO.fileExists(sFile) then FSO.CopyFile sfile, dfile 
	End Function
	
	' 删除文件
	' 参数 目标文件
	' 返回 无
	' 例 call DelFile("c:\abd\123.txt")	
	Public Function DelFile(sfile)
		if FSO.FileExists(sfile) then FSO.DeleteFile sfile,true
	End Function
	
	' 删除目录
	' 参数 目录
	' 返回 无
	' 例 call DelDir("c:\abd\")	
	Public Function DelDir(sPath)
		sPath=Replace(sPath,"/","\")
	    if Right(sPath,1)="\" then sPath=left(sPath,len(sPath)-1)
		if FSO.FolderExists(sPath) then FSO.DeleteFolder sPath
	End Function
	
	' 运行程序
	' 参数 程序 是否等待结束
	' 返回 无
	' 例 call Run("c:\abd\123.txt",false)	
	Public Function Run(sPath,wait)
	    if FSO.FileExists(sPath) then
			WshShell.run """"&sPath&"""",,wait
		end if
	End Function
	
	' ping机器是否在线
	' 参数 IP地址 
	' 返回true或false
	' 例 call ping("192.168.0.1")
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
	
	' 取得网卡MAC地址
	' 参数 无
	' 返回本机mac地址
	' 例 call GetMac	
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
	
	' 取得本机IP地址
	' 参数 无
	' 返回本机IP地址
	' 例 call GetIP
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

	' 取得机器名称
	' 参数 无
	' 返回本机机器名称
	' 例 call GetComputerName	
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
	
	' 取得操作系统名
	' 参数 无
	' 返回  操作系统名
	' 例 call GetOS	
	Public Function GetOs
	   ComputerName="."
		Dim objWMIService,colItems,objItem,objAddress
		Set objWMIService = GetObject("winmgmts:\\" & ComputerName & "\root\cimv2")
		Set colItems = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
		For Each objItem in colItems
			'GetOs = objItem.Caption&" 版本"& objItem.Version
			if instr(objItem.Version,"6.1")>0 then '6.0是vista 6.1是win7 6.2是win8 10.0是win10
			  GetOS="Win7"
			  exit for
			elseif instr(objItem.Version,"10.0")>0 then
			  GetOs="Win10"
			  exit for
			end if			
		Next	
	End Function
	
	' 取得 操作系统位数
	' 参数 无
	' 返回  操作系统位数 64位系统返回x64 32位系统返回x86
	' 例 call X86orX64	
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
	
	' 文件转成16进制字符串
	' 参数 文件名 16进制文件 如何第二个参数为空，直接返回16进制字符串
	' 返回16进制字符串 或存为文件    16进制文本文件会比可执行程序大一倍
	' 例生成字符串 call ReadBinary("c:\windows\notepad.exe","")
	' 例生成文本文件 call ReadBinary("c:\windows\notepad.exe","d:\123.txt")
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
	
	' 16进制字符串转成可执行文件 
	' 参数 字符串 可执行文件(完全路径) 是否是文件 
	' 返回 无
	' 例 字符串生成 call BinaryToFile("d:\123.exe","4D5A90000300000004000000FFFF",false)
	' 例 文本文件生成 call BinaryToFile("d:\123.exe","d:\123.txt",true)
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

	' 下载远程文件到本地
	' 参数 远程地址 本地文件
	' 返回 无
	' 例 call DownFile("https://dl.360safe.com/360sd/360sd_x64_std_5.0.0.8183C.exe","d:\360sd.exe")
	Public Function DownFile(UrlFile,SaveFile)
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
	
	' '延时函数	
	' 参数  秒
	' 返回 无
	' 例 call Sleep(5)
	Public Sub Sleep(sec)
		WScript.sleep sec*1000 
	End sub
	
	' 导入注册表文件
	' 参数 文件名
	' 返回 无
	' 例 call ImportReg("d:\1.reg")
	Public Function ImportReg(regFile)
	    if FSO.FileExists(regFile) then
			WshShell.run "regedit.exe /s """&regFile&"""",0
		end if
	End Function	
	
	' 运行bat文件
	' 参数 文件名
	' 返回 无
	' 例 Call RunBat(batFile)
	Public Function RunBat(batFile)
	    if FSO.FileExists(batFile) then
			WshShell.run """"&batFile&"""",0
		end if
	End Function

    ' 导入vbs文件 
    ' 参数 vbs文件
    ' 返回 无
    ' 例 call import("d:\abc.vbs")
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
	
	' 关闭指定进程 
	' 参数 进程名
	' 返回 无
	' 例 call CloseProcess("winrar.exe")
	Public Sub CloseProcess(ExeName)
	    if IsProcess(ExeName) then
		  WshShell.run "Taskkill /f /im " & ExeName,0
		end if
	End Sub

	' '检测进程  
	' 参数 进程名
	' 返回 进程正在运行，返回true
	' 例 Call IsProcess("qq.exe")	
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
	
	' '检测进程组
	' 参数 进程列表，进程之间用|分隔
	' 返回 进程列表中只要有一个进程在运行返回true
	' 例	Call IsProcessEx("qq.exe|notepad.exe")
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
	
	' '结束进程组
	' 参数 进程列表，中间用|分隔
	' 返回 无
	' 例	call CloseProcessEx("qq.exe｜wecat.exe")
	Public Sub CloseProcessEx(ExeName)
		dim ProcessName,CmdCode,i
		ProcessName = Split(ExeName, "|")
		For i=0 to UBound(ProcessName)
		    if IsProcess(ProcessName(i)) then  '如果进程存在
			  CmdCode=CmdCode & " /im " & ProcessName(i)
			  WshShell.run "Taskkill /f" & CmdCode,0
			end if
		Next		
	End Sub	
	
	' 正则匹配
	
	Public Function RegExpTest(patrn, strng)  
	  Set re = New RegExp  
	  re.Pattern = patrn  
	  re.IgnoreCase = True 
	  re.Global = True 
	  Set Matches = re.Execute(strng)  
	  RegExpTest = Matches.Count  
	  Set re=Nothing
	End Function
	
	' '写注册表
	' 参数 key 值 类型
	' 返回 无
	' 例	call WriteReg("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\MAIN\Start Page","https://www.baidu.com","")
	Public Sub WriteReg(regkey, value, typeName) 
		If typeName = "" Then
			WshShell.RegWrite regkey, value
		Else
			WshShell.RegWrite regkey, value, typeName
		End If
	End Sub

	' '读取注册表，搜索key，返回所在路径
	' 参数 key
	' 返回 无
	' 例	call ReadReg("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\MAIN\Start Page")
	Public Function ReadReg(regkey) '
	    on error resume next
		err.clear
		ReadReg = WshShell.RegRead(regkey)
		if err.number<>0 then
		  ReadReg=false
		 end if 
	End Function

	' '关闭指定标题窗口
	' 参数 类名 窗口名
	' 返回 无
	' 例	call KillWindow("","无标题")
	Public Function KillWindow(classname,winName)
		if len(classname)=0 then classname=0
		if len(winName)=0 then winName=0
		hwnd=DWX.FindWindow(classname,winName)
		if hwnd<>0 then
		  DWX.SendMessage hwnd,&H10,0,0 '关闭窗口
		'DWX.PostMessage hwnd,&H112,&HF060, 0 '关闭窗口
		'DWX.PostMessage hwnd, &H82, 0, 0 '销毁窗口
		end if
	   'dim rcSuccess  '使用wscript发送alt+F4
	   'rcSuccess = WshShell.AppActivate(winName)
	   'if rcSuccess then WshShell.sendkeys "%{F4}"
	End Function
	
	' '隐藏指定标题窗口
	' 参数 类名 窗口名
	' 返回 无
	' 例	call HideWindow("Notepad","")
	Public Function HideWindow(classname,winName)
		if len(classname)=0 then classname=0
		if len(winName)=0 then winName=0
		hwnd=DWX.FindWindow(classname,winName)
		if hwnd<>0 then
	      hrgn =DWX.CreateRectRgn(0,0,0,0)
	      DWX.SetWindowRgn hwnd,hrgn,true '隐藏视界
		  DWX.ShowWindow hwnd,0  '隐藏窗口
		End If
	End Function

	' '按照窗口中止线程
	' 参数 类名 窗口名
	' 返回 无
	' 例	call KillThread("Notepad","")
	Public Function KillThread(classname,winName)
		if len(classname)=0 then classname=0
		if len(winName)=0 then winName=0
		hwnd=DWX.FindWindow(classname,winName)
		if hwnd<>0 then
			tid=DWX.GetWindowThreadProcessId(hwnd,0) '取得线程ID
			DWX.PostThreadMessage tid,&H12,0,0  '退出线程 
		end if
	End Function	
	
	' 同步网络时间
	' 参数 无
	' 返回 
	' 例 call SyncTime
	Public Sub SyncTime()
        On error resume next	
	    url = "http://free.timeanddate.com/clock/i1jyoa52/n236/tt0/tw0/tm3/td2/th1/tb4" 
		'url = "http://api.m.taobao.com/rest/api3.do?api=mtop.common.getTimestamp" '用淘宝时间api
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
		   WshShell.Run "cmd.exe /c date " & dtmNewDate,0 
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
		   WshShell.Run "cmd.exe /c time " & dtmNewTime,0 
		End if
	End Sub
	
	' 取得系统开机时间 关机时间 返回时间较长，可用于学习查询系统日志
	' 参数 无
	' 返回 
	' 例 call GetSysRunTime
	Public Function GetSysRunTime()
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
			GetSysRunTime= "本次开机时间: " & FormatWMIUTC(objEvent.TimeWritten)
		  ElseIf Flag = 2 Then
			GetSysRunTime=GetSysRunTime& " 上次关机时间: " & FormatWMIUTC(objEvent.TimeWritten)
		  ElseIf Flag = 3 Then
			GetSysRunTime=GetSysRunTime& " 上次开机时间: " & FormatWMIUTC(objEvent.TimeWritten)
			Exit For
		  End If
		Next
    End Function
	
	' 格式化wmi时间
	'FormatUTC
	Public Function FormatWMIUTC(WMIDateString)
	  DS = " // :: "
	  FormatWMIUTC = Left(WMIDateString,2)
	  For i = 2 To 7
		FormatWMIUTC = FormatWMIUTC & Mid(WMIDateString, i * 2 - 1, 2) & Mid(DS,i,1)
	  Next
	  'FormatWMIUTC = Mid(WMIDateString, 1, 4) & "年" _
	  '      & Mid(WMIDateString, 5, 2) & "月" _
	  '      & Mid(WMIDateString, 7, 2) & "日 " _
	  '      & Mid (WMIDateString, 9, 2) & ":" _
	  '      & Mid(WMIDateString, 11, 2) & ":" _
	  '      & Mid(WMIDateString,13, 2)
	End Function
	
	'把标准时间转换为UNIX时间戳
	'参数：strTime:要转换的时间；intTimeZone：该时间对应的时区       
	'返回值：strTime相对于1970年1月1日午夜0点经过的秒数       
	'示例：ToUnixTime("2008-5-23 10:51:0", +8)，返回值为1211511060       
	Public Function ToUnixTime(strTime, intTimeZone)       
		If IsEmpty(strTime) or Not IsDate(strTime) Then strTime = Now       
		If IsEmpty(intTimeZone) or Not isNumeric(intTimeZone) Then intTimeZone = 0       
		ToUnixTime = DateAdd("h",-intTimeZone,strTime)       
		ToUnixTime = DateDiff("s","1970-1-1 0:0:0", ToUnixTime)       
	End Function      
		  
	'把UNIX时间戳转换为标准时间       
	'参数：intTime:要转换的UNIX时间戳；intTimeZone：该时间戳对应的时区       
	'返回值：intTime所代表的标准时间       
	'示例：FromUnixTime("1211511060", +8)，返回值2008-5-23 10:51:0       
	Public Function FromUnixTime(intTime, intTimeZone)       
		If IsEmpty(intTime) Or Not IsNumeric(intTime) Then      
			FromUnixTime = Now()       
		   Exit Function      
		End If      
		If IsEmpty(intTime) Or Not IsNumeric(intTimeZone) Then intTimeZone = 0       
		FromUnixTime = DateAdd("s", intTime, "1970-1-1 0:0:0")       
		FromUnixTime = DateAdd("h", intTimeZone, FromUnixTime)       
	End Function	
	
	' 生成日志
	' 参数 日志内容：
	' 返回值  无
	' 示例: log("新加日志")
	Public Function log(logstr)	  
	    logfile=CurrentPath&"\log-"&year(Now)&"-"&Month(Now)&"-"&day(Now)&".txt"
	    if fso.FileExists(logfile) then
           Set ObjLog = FSO.OpenTextFile(LogFile,8)		   
		else 
          Set ObjLog = FSO.CreateTextFile(logfile)
		end if
		ObjLog.Write vbCrLf&"["&Now&"]  日志内容："&logstr
		ObjLog.close
		set ObjLog=Nothing
	End Function
	
	'判断当前是否无盘超级用户
	'参数：无
	'返回值 true 或false
	'示例：IsSuperAdmin
	Public Function IsSuperAdmin()
		'[网维大师/绿化大师/信佑Win无盘]
		AdminValue=ReadReg("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\iCafe8\Admin")
		If AdminValue=1 then
		    IsSuperAdmin=true
			exit Function
		End IF
		'[易乐游]
		AdminValue=ReadReg("HKEY_LOCAL_MACHINE\SOFTWARE\EYOOCLIENTSTATUS\SuperLogin")
		If AdminValue=1 then
		    IsSuperAdmin=true
			exit Function
		End IF
		'[云更新]
		AdminValue=ReadReg("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\SuperAdmin")
		If AdminValue=1 then
		    IsSuperAdmin=true
			exit Function
		End IF
		'[方格子]
		AdminValue=ReadReg("HKEY_LOCAL_MACHINE\SYSTEM\nVos\SupperMode")
		If AdminValue=1 then
		    IsSuperAdmin=true
			exit Function
		End IF
		'[锐起]
		AdminValue=ReadReg("HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Richdisk\ClientFlag")
		If AdminValue=1 then
		    IsSuperAdmin=true
			exit Function
		End IF
        IsSuperAdmin=false
	End Function
	
	'功能：读取INI文件
	'参数： 节点,键名，默认值，ini文件
	'返回值:键值
	'示例：ReadIni("节点","键名","默认值","d:\123.ini")
	Public Function ReadIni(iSection,iKey,dValue,iFile)
	  if len(iSection)<>0 and len(iFile)<>0 then 
		  dim vStr
		  vStr=Space(255)
		  Call DWX.GetPrivateProfileString(iSection,iKey,dValue,vStr,255,iFile)
		  ReadIni=vStr
		  Set vStr=Nothing
      End If		  
	End Function
	
	'功能：写INI文件
	'参数：节点，键名，键值，ini文件
	'返回值:写入成功返回1 否则返回0
	'示例：WriteIni "节点","键名","值","d:\123.ini"
	Public Function WriteIni(iSection,iKey,iValue,iFile)
	  if len(iSection)<>0 and len(iKey)<>0 and len(iFile)<>0 then 
	     WriteIni=DWX.WritePrivateProfileString(iSection,iKey,iValue,iFile)
	  End if
	End Function
end class