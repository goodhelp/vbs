class vbsfun
	rem 类实例化时执行的代码
	Public WshShell,FSO
	private sub Class_Initialize()
		Set WshShell = WScript.CreateObject("WScript.Shell")
		Set FSO=CreateObject("Scripting.FileSystemObject")
	end sub

	rem 类销毁时执行的代码
	private sub class_terminate()
		Set WshShell=Nothing
		Set FSO=Nothing
	end sub
	
	Rem 在桌面创建一个快捷方式 
	rem 参数：快捷方式名称  程序地址 程序运行参数 图标地址 
	rem 返回 无
	rem 例 call MakeLink("罗技鼠标设置","G:\常用软件\罗技鼠标游戏驱动\Rungame.exe","","G:\常用软件\罗技鼠标游戏驱动\48731.ico")
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
	
	rem 收藏夹添加网址
	rem 参数:网址 快捷名称 是否创建在收藏夹栏
	rem 返回 无
	rem 例 call MakeUrl("http://www.bnwin.com","百脑问",true)	
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
	
	rem 修改主页
	rem 参数 网址
	rem 返回
	rem 例 call SetHomepage("https://www.baidu.com")
	Public Function SetHomepage(url)
		WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Start Page",url	
	End Function
	
	rem 根据exe取所在路径
	rem 参数 完全路径  
	rem 返回 路径
	rem 例 call GetExePath("CProgram FilesInternet Explorer\iexplore.exe")
	Public Function GetExePath(strFileName)
		strFileName=Replace(strFileName,"/","\")
		dim ipos
		ipos=InstrRev(strFileName,"\")
		GetExePath=left(strFileName,ipos)
	End Function

	rem 判断文件是否存在 
	rem 参数 文件地址  
	rem 返回 true或false
	rem 例 call IsExitFile("c:\abc.txt")
	Public Function IsExitFile(filespec)     
        If FSO.fileExists(filespec) Then         
			IsExitFile=True        
        Else
			IsExitFile=False 
        End If
	End Function 
	
	rem 判断目录是否存在 
	rem 参数 目录地址 是否创建  
	rem 返回 true或false
	rem 例 call IsExitDir("c:\abc",true)
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
	
	rem 创建多级目录
	rem 参数  路径 
	rem 返回 无
	rem 例  call MyCreateFolder("c:\ad\1233\dd")
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
	
	rem 拷贝目录
	rem 参数 源目录  目录目录  是否覆蓋
	rem 返回 拷贝的文件数
	rem 例 call XCopy("c:\123","d:\123",true)
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

	rem 复制文件
	rem 参数 源文件 目标文件  是否覆蓋
	rem 返回 无
	rem 例 call CopyFile("c:\abd\123.txt","d:\323\aaa.txt",true)	
	Public Function CopyFile(sfile,dfile,overwrite)
		if (overwrite and FSO.FileExists(dfile)) then FSO.DeleteFile dfile,true
		if Not FSO.FileExists(GetExePath(dfile)) then
		  MyCreateFolder(GetExePath(dfile))
		end if
		if FSO.fileExists(sFile) then FSO.CopyFile sfile, dfile 
	End Function
	
	rem 删除文件
	rem 参数 目标文件
	rem 返回 无
	rem 例 call DelFile("c:\abd\123.txt")	
	Public Function DelFile(sfile)
		if FSO.FileExists(sfile) then FSO.DeleteFile sfile,true
	End Function
	
	rem 删除目录
	rem 参数 目录
	rem 返回 无
	rem 例 call DelDir("c:\abd\")	
	Public Function DelDir(sPath)
		sPath=Replace(sPath,"/","\")
	    if Right(sPath,1)="\" then sPath=left(sPath,len(sPath)-1)
		if FSO.FolderExists(sPath) then FSO.DeleteFolder sPath
	End Function
	
	rem 运行程序
	rem 参数 程序 是否等待结束
	rem 返回 无
	rem 例 call Run("c:\abd\123.txt",false)	
	Public Function Run(sPath,wait)
	    if FSO.FileExists(sPath) then
			WshShell.run """"&sPath&"""",,wait
		end if
	End Function
	
	rem ping机器是否在线
	rem 参数 IP地址 
	rem 返回true或false
	rem 例 call ping("192.168.0.1")
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
	
	rem 取得网卡MAC地址
	rem 参数 无
	rem 返回本机mac地址
	rem 例 call GetMac	
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
	
	rem 取得本机IP地址
	rem 参数 无
	rem 返回本机IP地址
	rem 例 call GetIP
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

	rem 取得机器名称
	rem 参数 无
	rem 返回本机机器名称
	rem 例 call GetComputerName	
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
	
	rem 取得操作系统名
	rem 参数 无
	rem 返回  操作系统名
	rem 例 call GetOS	
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
	
	rem 取得 操作系统位数
	rem 参数 无
	rem 返回  操作系统位数 64位系统返回x64 32位系统返回x86
	rem 例 call X86orX64	
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
	
	rem 文件转成16进制字符串
	rem 参数 文件名 16进制文件 如何第二个参数为空，直接返回16进制字符串
	rem 返回16进制字符串 或存为文件    16进制文本文件会比可执行程序大一倍
	rem 例生成字符串 call ReadBinary("c:\windows\notepad.exe","")
	rem 例生成文本文件 call ReadBinary("c:\windows\notepad.exe","d:\123.txt")
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
	
	rem 16进制字符串转成可执行文件 
	rem 参数 字符串 可执行文件(完全路径) 是否是文件 
	rem 返回 无
	rem 例 字符串生成 call BinaryToFile("d:\123.exe","4D5A90000300000004000000FFFF",false)
	rem 例 文本文件生成 call BinaryToFile("d:\123.exe","d:\123.txt",true)
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

	
	rem '延时函数	
	rem 参数  秒
	rem 返回 无
	rem 例 call Sleep(5)
	Public Sub Sleep(sec)
		WScript.sleep sec*1000 
	End sub
	
	rem 导入注册表文件
	rem 参数 文件名
	rem 返回 无
	rem 例 call ImportReg("d:\1.reg")
	Public Function ImportReg(regFile)
	    if FSO.FileExists(regFile) then
			WshShell.run "regedit.exe /s """&regFile&"""",0
		end if
	End Function	
	
	rem 运行bat文件
	rem 参数 文件名
	rem 返回 无
	rem 例 Call RunBat(batFile)
	Public Function RunBat(batFile)
	    if FSO.FileExists(batFile) then
			WshShell.run """"&batFile&"""",0
		end if
	End Function

    rem 导入vbs文件 
    rem 参数 vbs文件
    rem 返回 无
    rem 例 call import("d:\abc.vbs")
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
	
	rem 关闭指定进程 
	rem 参数 进程名
	rem 返回 无
	rem 例 call CloseProcess("winrar.exe")
	Public Sub CloseProcess(ExeName)
		WshShell.run "Taskkill /f /im " & ExeName,0
	End Sub

	rem '检测进程  
	rem 参数 进程名
	rem 返回 进程正在运行，返回true
	rem 例 Call IsProcess("qq.exe")	
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
	
	rem '检测进程组
	rem 参数 进程列表，进程之间用|分隔
	rem 返回 进程列表中只要有一个进程在运行返回true
	rem 例	Call IsProcessEx("qq.exe|notepad.exe")
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
	
	rem '结束进程组
	rem 参数 进程列表，中间用|分隔
	rem 返回 无
	rem 例	call CloseProcessEx("qq.exe｜wecat.exe")
	Public Sub CloseProcessEx(ExeName)
		dim ProcessName,CmdCode,i
		ProcessName = Split(ExeName, "|")
		For i=0 to UBound(ProcessName)
			CmdCode=CmdCode & " /im " & ProcessName(i)
			WshShell.run "Taskkill /f" & CmdCode,0
		Next		
	End Sub	
	
	rem 正则匹配
	
	Public Function RegExpTest(patrn, strng)  
	  Set re = New RegExp  
	  re.Pattern = patrn  
	  re.IgnoreCase = True 
	  re.Global = True 
	  Set Matches = re.Execute(strng)  
	  RegExpTest = Matches.Count  
	End Function

end class