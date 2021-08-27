set myfun=New vbsfun
'call myfun.CopyFile("d:\Users\Administrator\Desktop\myvbsfun.vbs","d:\aa\123.vbs",true)
'call myfun.Run("regedit",true)
MsgBox("结束")

class vbsfun
	rem 类实例化时执行的代码
	private sub Class_Initialize()

	end sub

	rem 类销毁时执行的代码
	private sub class_terminate()

	end sub

	Rem 在桌面创建一个记事本快捷方式 
	rem 参数：快捷方式名称  程序地址 程序运行参数 图标地址 
	rem 返回 无
	rem 例 call MakeLink("罗技鼠标设置.lnk","G:\常用软件\罗技鼠标游戏驱动\Rungame.exe","","G:\常用软件\罗技鼠标游戏驱动\48731.ico")
	Public Function MakeLink(linkname,linkexe,linkparm,linkico)
		Set WshShell = WScript.CreateObject("WScript.Shell")
		strDesktop = WshShell.SpecialFolders("Desktop") rem 特殊文件夹“桌面”
		set oShellLink = WshShell.CreateShortcut(strDesktop &"\"& linkname)
		oShellLink.TargetPath = linkexe  '可执行文件路径
		oShellLink.Arguments = linkparm '程序的参数
		oShellLink.WindowStyle = 1 '参数1默认窗口激活，参数3最大化激活，参数7最小化
		oShellLink.Hotkey = ""  '快捷键
		if IsExitAFile(linkico) then
		oShellLink.IconLocation = linkico&", 0" '图标
		else
		oShellLink.IconLocation = "%SystemRoot%\system32\SHELL32.dll,8"
		end if
		oShellLink.Description = ""  '备注
		oShellLink.WorkingDirectory = GetExePath(linkexe)  '起始位置
		oShellLink.Save  '创建保存快捷方式	
		Set WshShell=Nothing
		Set oShellLink=Nothing
	End Function
	
	rem 根据exe取所在路径
	rem 参数 完全路径  
	rem 返回 路径
	rem 例 call GetExePat("CProgram FilesInternet Explorer\iexplore.exe")
	Public Function GetExePath(strFileName)
		strFileName=Replace(strFileName,"/","\")
		dim ipos
		ipos=InstrRev(strFileName,"\")
		GetExePath=left(strFileName,ipos)
	End Function

	rem 判断文件是否存在 
	rem 参数 文件地址  
	rem 返回 true或false
	rem 例 call IsExitAFile("c:\abc.txt")
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
	
	rem 判断目录是否存在 
	rem 参数 目录地址 是否创建  
	rem 返回 true或false
	rem 例 call IsExitDir("c:\abc",true)
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
	
	rem 创建多级目录
	rem 参数  路径 
	rem 返回 无
	rem 例  call MyCreateFolder("c:\ad\1233\dd")
	Public Sub MyCreateFolder(sPath)
		sPath=Replace(sPath,"/","\")
		if Right(sPath,1)="\" then sPath=left(sPath,len(sPath)-1) '删除目录末尾的\
		Dim fso
		set fso = CreateObject("Scripting.FileSystemObject")
		if(Len(sPath) > 0 And fso.FolderExists(sPath) = False) Then
			Dim pos, sLeft
			pos = InStrRev(sPath, "\")
			if(pos <> 0) Then
				sLeft = Left(sPath, pos - 1)
				MyCreateFolder sLeft            '先创建父目录
			end if
			fso.CreateFolder sPath              '再创建本目录
		end if
		set fso = Nothing
	End Sub
	
	rem 拷贝目录
	rem 参数 源目录  目录目录  是否覆蓋
	rem 返回 拷贝的文件数
	rem 例 call XCopy("c:\123" "d:\123",true)
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

	rem 复制文件
	rem 参数 源文件 目标文件  是否覆蓋
	rem 返回 无
	rem 例 call CopyFile("c:\abd\123.txt","d:\323\aaa.txt",true)	
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
	
	rem 删除文件
	rem 参数 目标文件
	rem 返回 无
	rem 例 call DelFile("c:\abd\123.txt")	
	Public Function DelFile(sfile)
		dim fso
		set fso = CreateObject("Scripting.FileSystemObject")
		if fso.FileExists(sfile) then fso.DeleteFile sfile,true
		set fso = nothing
	End Function
	
	rem 删除目录
	rem 参数 目录
	rem 返回 无
	rem 例 call DelDir("c:\abd\123.txt")	
	Public Function DelDir(sPath)
		sPath=Replace(sPath,"/","\")
	    if Right(sPath,1)="\" then sPath=left(sPath,len(sPath)-1)
		dim fso
		set fso = CreateObject("Scripting.FileSystemObject")
		if fso.FolderExists(sPath) then fso.DeleteFolder sPath
		set fso = nothing
	End Function
	
	rem 运行程序
	rem 参数 程序 是否等待结束
	rem 返回 无
	rem 例 call Run("c:\abd\123.txt",false)	
	Public Function Run(sPath,wait)
		dim shell
		Set shell = Wscript.createobject("WScript.shell")
		shell.run """"&sPath&"""",,wait
		set shell = nothing
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
	
end class