'myfun为实例化MyVbsClass.vbs，可以使用全部函数
while(true)
	call myfun.sleep(2)
	call myfun.CloseProcessEx("x-panda.exe|lol_monitor2.exe|pubg_monitor2.exe|khardware64_v54.exe") '结束进程
	call myfun.KillWindow("","计算机") '关闭窗口
	'call myfun.HideWindow("","") '隐藏窗口
	'call myfun.KillThread("","")  '按窗口中止线程
Wend