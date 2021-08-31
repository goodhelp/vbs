'myfun为实例化MyVbsClass.vbs，可以使用全部函数
dim slist1,runloop,slist2,slist3,slist4,slist
runloop=false
'runexe和inifile以及n在runpg.vbs已定义
'查杀进程
for n=1 to 20
   runexe=myfun.ReadIni("循环查杀",CStr(n),"",inifile)
   if len(runexe)=0 then
      exit for
   elseif n=1 then
      slist1=runexe
   else
      slist1=slist1&"|"&runexe
   end if
Next
if len(slist1)<>0 then runloop=true
'关闭窗口
for n=1 to 20
   runexe=myfun.ReadIni("关闭窗口",CStr(n),"",inifile)
   if len(runexe)=0 then
      exit for
   elseif n=1 then
      slist2=runexe
   else
      slist2=slist2&"|"&runexe
   end if
Next
if len(slist2)<>0 then runloop=true
'隐藏窗口
for n=1 to 20
   runexe=myfun.ReadIni("隐藏窗口",CStr(n),"",inifile)
   if len(runexe)=0 then
      exit for
   elseif n=1 then
      slist3=runexe
   else
      slist3=slist3&"|"&runexe
   end if
Next
if len(slist3)<>0 then runloop=true
'中止线程
for n=1 to 20
   runexe=myfun.ReadIni("中止线程",CStr(n),"",inifile)
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
	call myfun.CloseProcessEx(slist1) '结束进程
	if len(slist2)<>0 then
		slist=split(slist2,"|")
		for n=0 to UBound(slist)
		   call myfun.KillWindow("",slist(n)) '关闭窗口
		next
	end if
	if len(slist3)<>0 then
		slist=split(slist3,"|")
		for n=0 to UBound(slist)
		   call myfun.HideWindow("",slist(n)) '隐藏窗口
		next
	end if
	if len(slist4)<>0 then
		slist=split(slist4,"|")
		for n=0 to UBound(slist)
		   call myfun.KillThread("",slist(n)) '按窗口中止线程
		next
	end if
Wend
if runloop=false then call myfun.log("无循环任务")