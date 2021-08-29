Set DWX = CreateObject("DynamicWrapperX")
DWX.Register "user32", "EnumWindows", "i=ph"
DWX.Register "user32", "GetWindowTextW", "i=hpl" ' variante Unicode
'DWX.Register "user32","GetWindowText","i=hpl"   ' variante ANSI

Set Ref = GetRef("CbkEnumWin")  

pCbkFunc = DWX.RegisterCallback(Ref, "i=hh", "r=l") ' enregistrement CbkEnumWin
													
                                                    
                                                     
n = 0 : m = 0 : WinList = ""
Buf = DWX.MemAlloc(256)               

DWX.EnumWindows pCbkFunc, 0           
                               	      
DWX.MemFree Buf

WScript.Echo "Windows in total :" & m & vbCrLf & " With a title : " & n & _
              vbCrLf & vbCrLf & WinList


' la proc¨¦dure de rappel (callback)

Function CbkEnumWin(hwnd, lparam)
  DWX.GetWindowTextW hwnd , Buf , 128  
  Title = DWX.StrGet(Buf, "w")
  ' DWX.GetWindowText hwnd, Buf, 256   
  ' Title = DWX.StrGet(Buf, "s")
  If Len(Titre) > 0 Then         	   
    WinList = WinList & hwnd & vbTab & Titre & vbCrLf
    n = n + 1
  End If
  m = m + 1
  CbkEnumWin = 1  
End Function

'https://omen999.developpez.com/tutoriels/vbs/dynawrapperx-v2-1/