'myfunΪʵ����MyVbsClass.vbs������ʹ��ȫ������
dim n,runexe,inifile
inifile=vbsPath&"\default.ini"
for n=1 to 20
    runexe=myfun.ReadIni("���г���",CStr(n),"",inifile)
	if len(runexe)=0 then exit for
	if IsNumeric(runexe) then	
	   call myfun.sleep(runexe)
	else
	   call myfun.Run(runexe,false) 
	end if
Next
call myfun.log("����ִ�����")