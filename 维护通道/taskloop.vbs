'myfunΪʵ����MyVbsClass.vbs������ʹ��ȫ������
dim slist1,runloop,slist2,slist3,slist4,slist
runloop=false
'runexe��inifile�Լ�n��runpg.vbs�Ѷ���
'��ɱ����
for n=1 to 20
   runexe=myfun.ReadIni("ѭ����ɱ",CStr(n),"",inifile)
   if len(runexe)=0 then
      exit for
   elseif n=1 then
      slist1=runexe
   else
      slist1=slist1&"|"&runexe
   end if
Next
if len(slist1)<>0 then runloop=true
'�رմ���
for n=1 to 20
   runexe=myfun.ReadIni("�رմ���",CStr(n),"",inifile)
   if len(runexe)=0 then
      exit for
   elseif n=1 then
      slist2=runexe
   else
      slist2=slist2&"|"&runexe
   end if
Next
if len(slist2)<>0 then runloop=true
'���ش���
for n=1 to 20
   runexe=myfun.ReadIni("���ش���",CStr(n),"",inifile)
   if len(runexe)=0 then
      exit for
   elseif n=1 then
      slist3=runexe
   else
      slist3=slist3&"|"&runexe
   end if
Next
if len(slist3)<>0 then runloop=true
'��ֹ�߳�
for n=1 to 20
   runexe=myfun.ReadIni("��ֹ�߳�",CStr(n),"",inifile)
   if len(runexe)=0 then
      exit for
   elseif n=1 then
      slist4=runexe
   else
      slist4=slist4&"|"&runexe
   end if
Next
if len(slist4)<>0 then runloop=true
while(runloop)
	call myfun.sleep(2)
	call myfun.CloseProcessEx(slist1) '��������
	if len(slist2)<>0 then
		slist=split(slist2,"|")
		for n=0 to UBound(slist)
		   call myfun.KillWindow("",slist(n)) '�رմ���
		next
	end if
	if len(slist3)<>0 then
		slist=split(slist3,"|")
		for n=0 to UBound(slist)
		   call myfun.HideWindow("",slist(n)) '���ش���
		next
	end if
	if len(slist4)<>0 then
		slist=split(slist4,"|")
		for n=0 to UBound(slist)
		   call myfun.KillThread("",slist(n)) '��������ֹ�߳�
		next
	end if
Wend
if runloop=false then call myfun.log("��ѭ������")