'myfun为实例化MyVbsClass.vbs，可以使用全部函数
dim n,runexe
'inifile为main.vbs中定义的
for n=1 to 20
    runexe=myfun.ReadIni("运行程序",CStr(n),"",inifile)
	if len(runexe)=0 then exit for
	if IsNumeric(runexe) then	
	   call myfun.sleep(runexe)
	else
	   call myfun.Run(runexe,false) 
	end if
Next
call myfun.log("程序执行完成")