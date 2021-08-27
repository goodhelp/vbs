import("E:\Èí¼þ¹¤³Ì\vbs\myvbsfun.vbs")
set myfun=New vbsfun
'call myfun.CopyFile("d:\Users\Administrator\Desktop\myvbsfun.vbs","d:\aa\123.vbs",true)
'call myfun.Run("regedit",true)
'call myfun.SetHomepage("http://www.bnwin.com")
'call myfun.ReadBinary("c:\windows\notepad.exe","d:\123.txt")
'call myfun.BinaryToFile("d:\123.txt","d:\123.exe",true)
'call myfun.RunBat(".\test.bat")
'call myfun.ImportReg(".\reg.reg")
'call myfun.CloseProcess("SunloginRemote.exe")
'MsgBox(myfun.IsProcess("qq.exe"))
MsgBox(myfun.GetComputerName)

dim cptname
cptname=myfun.GetComputerName


set myfun=nothing

Sub import(sFile)
	Dim oFSO, oFile
	Dim sCode
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