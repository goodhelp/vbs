Set DWX = CreateObject("DynamicWrapperX")
DWX.Register "user32", "EnumWindows", "i=ph"
DWX.Register "user32", "GetWindowTextW", "i=hpl" ' variante Unicode
'DWX.Register "user32","GetWindowText","i=hpl"   ' variante ANSI
DWX.Register "user32", "GetClassName", "i=hpl"

Set Ref = GetRef("CbkEnumWin")  
pCbkFunc = DWX.RegisterCallback(Ref, "i=hh", "r=l") ' enregistrement CbkEnumWin                                            
n = 0 : m = 0 : WinList = ""
Buf = DWX.MemAlloc(256)
Buf1=DWX.MemAlloc(256)               
DWX.EnumWindows pCbkFunc, 0                             	      
DWX.MemFree Buf
DWX.MemFree Buf1
WScript.Echo "窗口总数 :" & m & vbCrLf & " 其中有标题窗口 : " & n & vbCrLf & vbCrLf & WinList

Function CbkEnumWin(hwnd, lparam)
  DWX.GetWindowTextW hwnd , Buf , 128
  DWX.GetClassName hwnd,Buf1,256  
  Title = DWX.StrGet(Buf, "w")
  ClassName = DWX.StrGet(Buf1,"s")
  ' DWX.GetWindowText hwnd, Buf, 256   
  ' Title = DWX.StrGet(Buf, "s")
  If Len(Title) > 0 Then         	   
    WinList = WinList & hwnd & vbTab & ClassName & vbTab & Title & vbCrLf '窗口句柄  类名 标题
    n = n + 1
  End If
  m = m + 1
  CbkEnumWin = 1  
End Function

'https://omen999.developpez.com/tutoriels/vbs/dynawrapperx-v2-1/