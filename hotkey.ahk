#include function.ahk

;=====================================================================o
;                         系统设置                                     |
;=====================================================================o

;新建空白markdown文档
#+m::
{
    filePath := GetNewFilePath()
    filePath .= ".md"
    FileAppend("", filePath)
}
;新建空白txt文档
#+i::
{
    filePath := GetNewFilePath()
    filePath .= ".txt"
    FileAppend("", filePath)
}
;新建空白无后缀文件
#+u::
{
    filePath := GetNewFilePath()
    FileAppend("", filePath)
}

;关闭显示器

#+l::
{
  Run("nircmd.exe monitor off")
}

;=====================================================================o
;                         快速启动                                     |
;=====================================================================o

;打开便签
#+y::
{
  Run("C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Stickynote.lnk")
}

;=====================================================================o
;                         电源管理                                     |
;=====================================================================o

;切换到节电模式
#+1::
{
	exitCode := RunWait("C:\Windows\System32\powercfg.exe setactive 6ed08f3e-52a1-4d96-9ae0-2f6619c8cdfd", ,"Hide")
  if !exitCode
    SetWindowsInfo("电源计划切换","切换至节电模式",-3000)
  else
    SetWindowsError("电源计划切换","切换节电模式失败")
}

;切换到平衡模式
#+2::
{
	exitCode := RunWait("C:\Windows\System32\powercfg.exe setactive 381b4222-f694-41f0-9685-ff5bb260df2e", ,"Hide")
  if !exitCode 
    SetWindowsInfo("电源计划切换","切换至平衡模式",-3000)
  else
    SetWindowsError("电源计划切换","切换平衡模式失败")
}

;切换到高性能模式
#+3::
{
	exitCode := RunWait("C:\Windows\System32\powercfg.exe setactive 7e8f1757-922c-46b1-86fe-e71b27942aa0", , "Hide")
  if !exitCode
    SetWindowsInfo("电源计划切换","切换至高性能模式",-3000)
  else
    SetWindowsError("电源计划切换","切换高性能模式失败")
}

;=====================================================================o
;                         音量管理                                     |
;=====================================================================o

;打开音量管理器，然后再摁一次则关闭管理器
#+v::
{
  if WinExist("ahk_exe SndVol.exe")
  {
    WinClose("ahk_exe SndVol.exe")
  }
  else
  {
    Run("C:\Windows\System32\SndVol.exe")
    ;一定要等待窗口开启，润命令只会打开不会等待打开完成
    WinWait("ahk_exe SndVol.exe")
    WinActivate("ahk_exe SndVol.exe")
  }
}

/*
nircmd的命令
changeappvolume [Process] [volume level] {Device Name/Index}
showsounddevices 可以输出所有的音频输出设备
process 可以指定exe文件名（chrome.exe）或者完整的路径文件名（C:\chrome.exe）
volume level参数是介于 0 和 1 之间的正数或负数。正数增加音量，负数减小音量。例如，如果要将音量从 20%（当前音量）增加到 70%，则应将此参数设置为 0.5
device name 不指定则使用默认输出设备。可以将设备索引指定为数值（0=第一个设备）也可以指定完整的设备名称（扬声器，耳机）
*/

/*
setappvolume
*/

/*
muteappvolume <process> <mute>
process 进程名称（chrome.exe）没有ahk_exe
mute 是一个布尔值， 1（静音），0（取消静音）
*/

;需要一个窗口来显示当前的程序音量
/*
;增加当前活动窗口的音量
#+up::
{
  activeExe := WinGetProcessName("A")
  volume := 0.05
  command := Format("nircmd.exe changeappvolume {1} {2}", activeExe , volume)

}

;减少当前活动窗口的音量
#+down::
{
  activeExe := WinGetProcessName("A")
  volume := -0.05
  command := Format("nircmd.exe changeappvolume {1} {2}", activeExe , volume)
}

*/

;禁音当前活动窗口的音量
#+p::
{
  activeExe := WinGetProcessName("A")
  mute := 1

  command := Format("nircmd.exe muteappvolume {1} {2}", activeExe , mute)
  exitCode := RunWait(command,,"Hide")
  if exitCode
    SetWindowsError("音量管理","设置程序静音失败")
}

;接触活动程序的静音，临时方案，将来要实现检测状态然后切换状态
#+o::
{
  activeExe := WinGetProcessName("A")
  mute := 0

  command := Format("nircmd.exe muteappvolume {1} {2}", activeExe , mute)
  exitCode := RunWait(command,,"Hide")
  if exitCode
    SetWindowsError("音量管理","设置程序取消静音失败")
}

;=====================================================================o
;                         自建工具                                     |
;=====================================================================o

;剪贴板操作资料：https://blog.51cto.com/u_15127700/4163445
;剪贴板图片函数：https://github.com/wanglong001/ClipMd

;=====================================================================o
;                         笔记本宏                                     |
;=====================================================================o

;控制媒体暂停
AppsKey & space::
{
  Send ("{Media_Play_Pause}")
}

;控制媒体播放上一个内容
AppsKey & Left::
{
  send("{Media_Prev}")
}

;控制媒体播放下一个内容
AppsKey & Right::
{
  send("{Media_Next}")
}

