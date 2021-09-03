Option Explicit
'************************************
'* Sample GUI only with API calls   *
'* Need DynaWrap component          *
'* Use Struct v1.1 Class            *
'* syntax Win NT et >               *
'* omen999 february 2007            *
'************************************

Class Struct ' v1.1  allow typedef with dynawrap calls
  Public Property Get Ptr '******************************* Property Ptr
    Ptr=GetBSTRPtr(sBuf)
  End Property
  Public Sub Add(sItem,sType,Data) '********************** Method Add
    Dim lVSize,iA,iB,iD
    iA=InStr(1,sType,"[",1)
    iB=InStr(1,sType,"]",1)
    iD="0"
    If iA>0  And iB>0 Then
      iD=Mid(sType,iA+1,iB-iA-1)
      If isNumeric(iD) Then
        sType=Left(sType,iA-1)
      Else
        Err.raise 10000,"Method Add","The index " & iD & " must be numeric"
        Exit Sub
      End If
    End If
    Select Case UCase(sType)'************************************************* COMPLETE WITH OTHERS WIN32 TYPES
    'OS 32bits...
    Case "DWORD","LONG","WPARAM","LPARAM","POINTX","POINTY","ULONG","HANDLE","HWND","HINSTANCE","HDC","WNDPROC","HICON","HCURSOR","HBRUSH"
      lVSize=4
    Case "LPBYTE","LPCTSTR","LPSTR","LPPRINTHOOKPROC","LPSETUPHOOKPROC","LPVOID","INT","UINT"
      lVSize=4
    Case "WORD"
      lVSize=2
    Case "BYTE"
      lVSize=1
    Case "TCHAR"
      If CLng(iD)<1 Then lVSize="254" Else lVSize=iD
    Case Else
      Err.raise 10000,"Method Add","The type " & sType & " is not a Win32 type."
      Exit Sub
    End Select
    dBuf.Add sItem,lVSize
    sBuf=sBuf & String(lVSize/2+1,Chr(0))
    SetDataBSTR GetBSTRPtr(sBuf),lVSize,Data,iOffset
  End Sub
  Public Function GetItem(sItem) '********************************************** M¨¦thode GetItem
    Dim lOf,lSi,aItems,aKeys,i
    If dBuf.Exists(sItem) then
      lSi=CLng(dBuf.Item(sItem))
      aKeys=dBuf.Keys
      aItems=dBuf.Items
      lOf=0
      For i=0  To dBuf.Count-1
        If aKeys(i)=sItem Then Exit For
        lOf=lOf+aItems(i)
      Next
      GetItem=GetDataBSTR(Ptr,lSi,lOf)
    Else
      GetItem=""
      err.raise 10000,"Method GetItem","The item " & sItem & " don't exist"
    End If
  End Function
  Public Function GetBSTRPtr(ByRef sData)
  'retun the TRUE address (variant long) of the sData string BSTR
    Dim pSource 
    Dim pDest
    If VarType(sData)<>vbString Then 'little check
      GetBSTRPtr=0
      err.raise 10000, "GetBSTRPtr", "The variable is not a string"
      Exit Function
    End If
    pSource=oSCat.lstrcat(sData,"")    'trick to return sData pointer
    pDest=oSCat.lstrcat(GetBSTRPtr,"")  'idem
    GetBSTRPtr=CLng(0)            'cast  function variable
    'l'adresse du contenu r¨¦el de sBuf (4octets) ¨¦crase le contenu de la variable GetBSTPtr  
    'les valeurs sont incr¨¦ment¨¦es de 8 octets pour  tenir compte du Type Descriptor
    oMM.RtlMovememory pDest+8,pSource+8,4 
  End Function
'**************************************************************************** IMPLEMENTATION
  Private oMM,oSCat,oAnWi 'objets wrapper API
  Private dBuf,sBuf,iOffset 
  Private  Sub Class_Initialize 'Constructeur
    Set oMM=CreateObject("DynamicWrapperX")
    oMM.Register "kernel32.dll","RtlMoveMemory","i=lll","r=l"
    Set oSCat=CreateObject("DynamicWrapperX")
    oSCat.Register "kernel32.dll","lstrcat","i=ws","r=l"    
    Set oAnWi=CreateObject("DynamicWrapperX")            
      oAnWi.Register "kernel32.dll","MultiByteToWideChar","i=llllll","r=l"
    Set dBuf=CreateObject("Scripting.Dictionary")
    sBuf=""
    iOffset=0
  End Sub  
  Private Sub SetDataBSTR(lpData,iSize,Data,ByRef iOfs)
  'Place une valeur Data de taille iSize ¨¤ l'adresse lpData+iOfs
    Dim lW,hW,xBuf
    Select Case iSize   'on commence par formater les valeurs num¨¦riques
    Case 1
      lW=Data mod 256   'formatage 8 bits
      xBuf=ChrB(lW)
    Case 2           'if any
      lW=Data mod 65536 'formatage 16 bits
      xBuf=ChrW(lW)    'formatage little-endian
    Case 4
      hW=Fix(Data/65536)'high word
      lW=Data mod 65536 'low word
      xBuf=ChrW(lW) & ChrW(hW) 'formatage little-endian
    Case Else        'bytes array, size iSize
      xBuf=Data
    End Select
    oMM.RtlMovememory lpData+iOfs,GetBSTRPtr(xBuf),iSize
    iOfs=iOfs+iSize 'maj l'offset
  End Sub
  Private Function GetDataBSTR(lpData,iSize,iOffset)
  'Read an iSize data to lpData+iOffset address
    Const CP_ACP=0       'code ANSI  
    Dim pDest,tdOffset
    'valeurs pour les donn¨¦es num¨¦riques
    pDest=oSCat.lstrcat(GetDataBSTR,"")
    tdOffset=8
    Select Case iSize ' cast de la variable fonction
    Case 1
      GetDataBSTR=CByte(0)
    Case 2
      GetDataBSTR=CInt(0)
    Case 4
      GetDataBSTR=CLng(0)
    Case Else  'a little bit more complicated with string data...
        GetDataBSTR=String(iSize/2,Chr(0))
        'la chaine variant BSTR stocke ses donn¨¦es ailleurs
      pDest=GetBSTRPtr(GetDataBSTR)
      tdOffset=0
    End Select
    'le contenu de la structure ¨¤ l'offset iOffset ¨¦crase le contenu de la variable GetDataBSTR (tenir compte du TD)
    oMM.RtlMovememory pDest+tdOffset,lpData+iOffset,iSize 
    if tdOffset=0 Then
      oAnWi.MultiByteToWideChar CP_ACP,0,lpData+iOffset,-1,pDest,iSize 'don't forget conversion Ansi->Wide
      GetDataBSTR=Replace(GetDataBSTR,Chr(0),"")                 'clean the trailer
    End If
  End Function 
End Class

Class XGui 'v1.0
' this class create a dialogbox only by api calls 
' it uses automation component DynaWrap and the struct class upper to allow typedef with dynawrap calls
' 4 public methods: CreateForm, ShowForm, RunForm et AddControl
' 1 public object dictionnary dFrmData which keys are name controls and stores data controls
' edit, static et button controls return content, listbox/combobox the selected item if exists, or empty string
' radiobutton and checkbox return true if checked or false
' groupbox control always return false
' each control must have unique name
' if the last letter of a checkbox ou radiobutton control name is "k", the control wil be checked
' close form without dictionnary data with esc key, Alt+F4, close button and system menu
' button controls haven't default behavior et must be manage by RunForm method
' this release 1.0  manages only "&ok" et "&cancel" buttons
' button ok closes the form and set data dictionnary, button cancel acts like esc key


Public dFrmData ' object dictionnary
Public Sub CreateForm(sCaption,lLeft,lTop,lWidth,lHeight,bOnTaskBar)
'Create a modeless invisible form
'sCaption: form caption
'lLeft,lTop: coordinates form
'lWidth, lHeight: form dimensions
'bOnTaskBar: if true (-1) form is display on taskbar
'no return value

  Const WS_VISIBLE=&H10000000
  Const WS_POPUP=&H80000000
  Const WS_OVERLAPPEDWINDOW=&HCF0000
  Dim hTask,fChild
  If bOnTaskBar Then
    hTask=0
    fChild=0
  Else
    hTask=hWsh
    fChild=WS_CHILD
  End If
  hWF=oWGui.CreateWindowExA(0,"#32770",sCaption&"",WS_OVERLAPPEDWINDOW+WS_POPUP+fChild,lLeft,lTop,lWidth,lHeight,hTask,0,hIns,0)
End Sub
Public Sub ShowForm(bAlwaysOnTop)
'display the form created by CreateForm
'bAlwaysOnTop: if true (-1) form always on top
'no return value

  Const HWND_TOP=0
  Const HWND_TOPMOST=-1
  Const SWP_SHOWWINDOW=&H40
  Const SWP_NOMOVE=&H2
  Const SWP_NOSIZE=&H1
  Dim fTop
  
  If bAlwaysOnTop Then fTop=HWND_TOPMOST Else fTop=HWND_TOP
  oWGui.SetWindowPos hWF,fTop,0,0,0,0,SWP_SHOWWINDOW+SWP_NOMOVE+SWP_NOSIZE
End Sub
Public Sub RunForm()
'form messages pump and dictionnary gestion
'no return value

  Const WM_COMMAND=&H111
  Const WM_SYSCOMMAND=&H112
  Const WM_KEYUP=&H101
  Const WM_LBUTTONUP=&H202
  Const GCW_ATOM=-32
  Const LB_GETCURSEL=&H188
  Const LB_ERR=-1
  Const LB_GETTEXT=&H189
  Const LB_GETTEXTLEN=&H18A
  Const GWL_STYLE=-16
  Const WS_CHILD=&H40000000
  Const WS_VISIBLE=&H10000000
  Const WS_TABSTOP=&H10000
  Const BS_AUTOCHECKBOX=&H3
  Const BS_AUTORADIOBUTTON=&H9
  Const BM_GETCHECK=&HF0
  Const BST_UNCHECKED=&H0
  Const BST_CHECKED=&H1
  Const BST_INDETERMINATE=&H2
  Const BST_PUSHED=&H4
  Const BST_FOCUS=&H8
  Const CP_ACP=0
  Const GWL_ID=-12
  Dim sCN,sCNW     'control content ansi/wide
  Dim aKData,aHData 'dictionnary contents keys/datas
  Dim lGetI       'index selected item (listbox)
  Dim lStyle       'button style
  Dim lKCode      'param message
  Dim n        'compteur
  
  Do While oWGui.GetMessageA(MSG.Ptr,hWF,0,0)>0 'Main loop messages pump
    If oWGui.IsDialogMessageA(hWF,MSG.ptr)<>0 Then
      Select Case MSG.GetItem("message")
      Case WM_KEYUP,WM_LBUTTONUP
        lKCode=MSG.GetItem("wParam")
        If MSG.GetItem("message")=WM_LBUTTONUP Then lKCode=13 'left mouse click -> enterkey
        Select Case lKCode 
        Case 27 'esc 
          dFrmData.RemoveAll
          oWGui.DestroyWindow hWF
          Exit Do
        Case 13,32 'enter or space when is an button control
          If oWGui.GetClassLongA(oWGui.GetFocus,GCW_ATOM)=49175 Then 'get atom button
            sCNW=UCase(GetBSTRCtrl(oWGui.GetFocus))
            If sCNW="&OK" Then   'it's ok button, so set dictionnary data and form close
              aKData=dFrmData.Keys   'control names array
              aHData=dFrmData.Items   'control handles array
              
              For n=0 To dFrmData.Count-1 'loop
                sCNW=""
                If oWGui.GetClassLongA(aHData(n),GCW_ATOM)=49178 Then 'get atom listbox
                  lGetI=oWGui.SendMessageA(aHData(n),LB_GETCURSEL,0,0)
                  If lGetI<>LB_ERR Then 'get the selected item if any
                    sCN=String(127,Chr(0))
                    sCNW=String(oWGui.SendMessageA(aHData(n),LB_GETTEXT,lGetI,MSG.GetBSTRPtr(sCN)),Chr(0))
                    oWaw.MultiByteToWideChar CP_ACP,0,MSG.GetBSTRPtr(sCN),-1,MSG.GetBSTRPtr(sCNW),LenB(sCNW)
                  End If
                Else
                  If oWGui.GetClassLongA(aHData(n),GCW_ATOM)=49175 Then 'get atom button
                    lStyle=oWGui.GetWindowLongA(aHData(n),GWL_STYLE)
                    If ((lStyle And BS_AUTOCHECKBOX)=BS_AUTOCHECKBOX) Or ((lStyle And BS_AUTORADIOBUTTON)=BS_AUTORADIOBUTTON) Then
                      sCNW=False
                      If oWGui.SendMessageA(aHData(n),BM_GETCHECK,0,0)=BST_CHECKED Then sCNW=True
                    Else 'other pushbouton
                      sCNW=GetBSTRCtrl(aHData(n))
                    End If
                  Else 'get data for edit, combo, static...
                    sCNW=GetBSTRCtrl(aHData(n))
                  End If
                End If
                dFrmData.Item(aKData(n))=sCNW 'la maj
              Next
              oWGui.DestroyWindow hWF
              Exit Do
            End If
            If sCNW="&ANNULER" Then
              dFrmData.RemoveAll
              oWGui.DestroyWindow hWF
              Exit Do
            End If  
          End If
        End Select
      Case WM_COMMAND,WM_SYSCOMMAND
        If (MSG.GetItem("wParam")=2) Or (MSG.GetItem("wParam")=61536) Then 'close button or system menu
          dFrmData.RemoveAll
          oWGui.DestroyWindow hWF
          Exit Do
        End If
      End Select
    Else
      oWGui.TranslateMessage MSG.Ptr
      oWGui.DispatchMessageA MSG.Ptr
    End If  
  Loop  
End Sub
Public Sub AddControl(sName,sClass,sData,lLeft,lTop,lWidth,lHeight)
'add a control on the form create by CreateForm method
'sName: unique control name
'sClass: one of global system class name
'sData: control data
'lLeft,lTop: control position on screen
'lWidth, lHeight: control dimensions
'no return value
  
  Const WS_EX_CLIENTEDGE=&H200
  Const DEFAULT_GUI_FONT=17
  Const WM_SETFONT=&H30
  Const WS_CHILD=&H40000000
  Const WS_VISIBLE=&H10000000
  Const WS_TABSTOP=&H10000
  Const GWL_ID=-12
  Const WS_VSCROLL=&H200000
  Const BS_AUTOCHECKBOX=&H3
  Const BS_AUTORADIOBUTTON=&H9
  Const BS_GROUPBOX=&H7
  Const BM_SETCHECK=&HF1
  Const BST_CHECKED=1
  Const LBS_HASSTRINGS=&H40
  Const CBS_DROPDOWN=&H2
  Const CB_ADDSTRING=&H143
  Const LB_ADDSTRING=&H180
  Const LBS_DISABLENOSCROLL=&H1000
  Dim hWn       'current control handle
  Dim sD        'current control data
  Dim cbBuf      'array list/combo data
  Dim sX        'types buttons
  Dim lStyle      'current control styles
  Dim lStyleEx    'extended styles 
  Dim lSL        'style liste or combo
  Dim fC        'flag check
  Dim fL        'flag list
  Dim n          'loop
  
  fC=False
  fL=False
  'parameters definition for CreateWindowEx according to class control
  Select Case UCase(sClass)
  Case "EDIT"
    sX=sClass
    sD=sData
    lStyle=WS_CHILD+WS_VISIBLE+WS_TABSTOP
    lStyleEx=WS_EX_CLIENTEDGE
  Case "STATIC"
    sX=sClass
    sD=sData
    lStyle=WS_CHILD+WS_VISIBLE
    lStyleEx=0
  Case "COMBOBOX"
    sX=sClass
    sD=""
    lStyle=WS_CHILD+WS_VISIBLE+CBS_DROPDOWN+WS_TABSTOP
    lStyleEx=0
    cbBuf=Split(sData,"|")
    fL=True    
    lSL=CB_ADDSTRING
  Case "LISTBOX"
    sX=sClass
    sD=""
    lStyle=WS_CHILD+WS_VISIBLE+WS_TABSTOP+WS_VSCROLL+LBS_HASSTRINGS+LBS_DISABLENOSCROLL
    lStyleEx=WS_EX_CLIENTEDGE
    cbBuf=Split(sData,"|")
    fL=True
    lSL=LB_ADDSTRING
  Case "BUTTON"
    sX=sClass
    sD=sData
    lStyle=WS_CHILD+WS_VISIBLE+WS_TABSTOP
    lStyleEx=0
  Case "GROUPBOX"
    sX="button"
    sD=sData
    lStyle=WS_CHILD+WS_VISIBLE+BS_GROUPBOX
    lStyleEx=0
  Case "CHECKBOX"
    sX="button"
    sD=sData
    lStyle=WS_CHILD+WS_VISIBLE+WS_TABSTOP+BS_AUTOCHECKBOX
    lStyleEx=0
    fC=True
  Case "RADIOBUTTON"
    sX="button"
    sD=sData
    lStyle=WS_CHILD+WS_VISIBLE+WS_TABSTOP+BS_AUTORADIOBUTTON
    lStyleEx=0
    fC=True
  Case Else
    Err.raise 10000,"Method AddControl","The class " & sClass & " is not a global system class"
    Exit Sub
  End Select
  hWn=oWGui.CreateWindowExA(lStyleEx,sX&"",sD&"",lStyle,lLeft,lTop,lWidth,lHeight,hWF,0,hIns,0) 'control creation
  oWGui.SendMessageA hWn,WM_SETFONT,oWGui.GetStockObject(DEFAULT_GUI_FONT),-1             'default font
  If fL Then 'feed the listbox/combobox
    For n=0 to UBound(cbBuf)
      oWsm.SendMessageA hWn,lSL,0,MSG.GetBSTRPtr(cbBuf(n))
    Next
  End If
  If fC Then 'check control with end's name is letter k
    If UCase(Right(sName,1))="K" Then oWGui.SendMessageA hWn,BM_SETCHECK,BST_CHECKED,0
  End If
  dFrmData.Add sName,hWn 'add control handle to dictionnary
End Sub
'************************************************************************************************************* IMPLEMENTATION
Private oWGui   'object API GUI
Private oWsm   'object SendMessage (syntax different)
Private oWaw  'object ANSI -> UNICODE conversion 

Private MSG     'structure MSG from API
Private hIns    'instance handle
Private hWsh    'main window WScript handle (hidden)
Private hWF      'form handle

Private  Sub Class_Initialize 'Constructor
  Const GWL_HINSTANCE=-6
  Set oWGui=CreateObject("DynamicWrapperX")
  Set oWsm=CreateObject("DynamicWrapperX")
  Set oWaw=CreateObject("DynamicWrapperX")
  With oWGui
    .Register "user32.dll","FindWindowA","f=s","i=ss","r=l"
    .Register "user32.dll","CreateWindowExA","f=s","i=lsslllllllll","r=l"
    .Register "user32.dll","SetWindowPos","f=s","i=lllllll","r=l"
    .Register "user32.dll","GetMessageA","f=s","i=llll","r=l"
    .Register "user32.dll","DispatchMessageA","f=s","i=l","r=l"
    .Register "user32.dll","TranslateMessage","i=l","f=s","r=l"
    .Register "user32.dll","GetWindowLongA","f=s","i=ll","r=l"
    .Register "user32.dll","SendMessageA","f=s","i=llll","r=l"
    .Register "user32.dll","SetWindowLongA","f=s","i=lll","r=l"
    .Register "user32.dll","GetWindowLongA","f=s","i=ll","r=l"
    .Register "user32.dll","IsDialogMessageA","f=s","i=ll","r=l"
    .Register "user32.dll","DestroyWindow","f=s","i=l","r=l"
    .Register "user32.dll","GetFocus","f=s","r=l"
    .Register "user32.dll","GetWindowTextA","f=s","i=lll","r=l"
    .Register "user32.dll","GetWindowTextLengthA","f=s","i=l","r=l"
    .Register "user32.dll","GetClassLongA","f=s","i=ll","r=l"
    .Register "gdi32.dll","GetStockObject","f=s","i=l","r=l"
  End With
  oWsm.Register "user32.dll","SendMessageA","f=s","i=llls","r=l" 'di
  oWaw.Register "kernel32.dll","MultiByteToWideChar","f=s","i=llllll","r=l"
  Set MSG=New Struct
  With MSG
    .Add "hwnd","HWND",0 
     .Add "message","UINT",0
     .Add "wParam","WPARAM",0
     .Add "lParam","LPARAM",0
     .Add "time","DWORD",0
     .Add "ptx","POINTX",0
     .Add "pty","POINTY",0
  End With
  Set dFrmData=CreateObject("Scripting.Dictionary")
  hWsh=oWGui.FindWindowA("WSH-Timer",chr(0))
  hIns=oWGui.GetWindowLongA(hWsh,GWL_HINSTANCE)  
End Sub
Private Function GetBSTRCtrl(hdW)
' Return handle hdW control content as string BSTR
  Const CP_ACP=0
  Dim sBuf,sBufW
  sBuf=String(oWGui.GetWindowTextLengthA(hdW),Chr(0))  
  sBufW=String(oWGui.GetWindowTextA(hdW,MSG.GetBSTRPtr(sBuf),oWGui.GetWindowTextLengthA(hdW)+1),Chr(0))
  oWaw.MultiByteToWideChar CP_ACP,0,MSG.GetBSTRPtr(sBuf),-1,MSG.GetBSTRPtr(sBufW),LenB(sBufW)
  GetBSTRCtrl=sBufW
End Function
End Class

'************************************************************************* DialogBox SAMPLE

Dim oFrm
Set oFrm=New XGui
oFrm.CreateForm "DialogBox by omen999",150,300,480,300,-1 ' modeless form
oFrm.AddControl "label1","static","&Last Name :",10,8,60,16
oFrm.AddControl "edit1","edit","",10,26,120,20
oFrm.AddControl "label2","static","&First Name :",10,50,60,16
oFrm.AddControl "edit2","edit","",10,68,120,20
oFrm.AddControl "label3","static","A&ddress :",10,94,100,16
oFrm.AddControl "edit3","edit","",10,112,150,20
oFrm.AddControl "label4","static","&City :",10,136,100,20
oFrm.AddControl "edit4","edit","",10,152,100,20
oFrm.AddControl "gbox1","groupbox"," Sex ",6,178,84,72
oFrm.AddControl "rdbox1","radiobutton","&Male",10,194,68,18
oFrm.AddControl "rdbox2k","radiobutton","&Female",10,212,68,18   'this control will be checked
oFrm.AddControl "rdbox3","radiobutton","&Don't know",10,230,74,18
oFrm.AddControl "label5","static","&Status :",146,8,40,16
oFrm.AddControl "cbox1","combobox","single|married|divorcee",146,26,150,80
oFrm.AddControl "label6","static","&Type :",310,8,40,16
oFrm.AddControl "lbox1","listbox","anorexic|very thin|thin|normal|fat|obese|dead",310,28,150,80
oFrm.AddControl "ckbox1k","checkbox","Mem&ber",310,90,68,20      'this control will be checked
oFrm.AddControl "button1","button","&OK",392,240,70,24
oFrm.AddControl "button2","button","&Cancel",312,240,70,24
oFrm.ShowForm False
oFrm.RunForm 'messages pump

'display the dialogbox final content
MsgBox oFrm.dFrmData.Item("label1") & vbLf &_
     oFrm.dFrmData.Item("edit1") & vbLf &_
     oFrm.dFrmData.Item("label2") & vbLf &_
     oFrm.dFrmData.Item("edit2") & vbLf &_
     oFrm.dFrmData.Item("label3") & vbLf &_
     oFrm.dFrmData.Item("edit3") & vbLf &_
     oFrm.dFrmData.Item("label4") & vbLf &_
     oFrm.dFrmData.Item("edit4") & vbLf &_
     oFrm.dFrmData.Item("gbox1") & vbLf &_
     oFrm.dFrmData.Item("rdbox1") & vbLf &_
     oFrm.dFrmData.Item("rdbox2k") & vbLf &_
     oFrm.dFrmData.Item("rdbox3") & vbLf &_
     oFrm.dFrmData.Item("label5") & vbLf &_
     oFrm.dFrmData.Item("cbox1") & vbLf &_
     oFrm.dFrmData.Item("label6") & vbLf &_
     oFrm.dFrmData.Item("lbox1") & vbLf &_
     oFrm.dFrmData.Item("ckbox1k") & vbLf &_
     oFrm.dFrmData.Item("button1") & vbLf &_
     oFrm.dFrmData.Item("button2")