;实现自定义函数和维护全局变量

;=====================================================================o
;                         全局变量                                     |
;=====================================================================o


;=====================================================================o
;                         全局函数                                     |
;=====================================================================o

;获取资源管理器当前显示的路径
GetObjDir()
{
  Process := WinGetProcessName("A")
  class := WinGetClass("A")
  ; 活动窗口必须为桌面或资源管理器, 否则显示错误!
  If (Process != "explorer.exe")
  {
    ;MsgBox("Error!")  ; 可自定义错误处理.
    Exit()
    MsgBox("Error!")
  }
  If (class ~= "rogman|WorkerW")
  {
    ObjDir := A_Desktop
  }
  Else If (class ~= "(Cabinet|Explore)WClass")
  {
    for window in ComObject("Shell.Application").Windows  ; 可以考虑从地址栏获取当前路径
      If (window.hwnd = WinExist("A"))
        ObjDir := window.LocationURL
    ; StrReplace() is not case sensitive
    ; check for StringCaseSense in v1 source script
    ; and change the CaseSense param in StrReplace() if necessary
    ObjDir := StrReplace(ObjDir, "file:///",,,, 1)
    While FoundPos := RegExMatch(ObjDir, "i)(?<=%)[\da-f]{1,2}", &hex)  ; 在路径中含特殊符号时还原这些符号
          ; StrReplace() is not case sensitive
          ; check for StringCaseSense in v1 source script
          ; and change the CaseSense param in StrReplace() if necessary
          ObjDir := StrReplace(ObjDir, "`%" hex[0], Chr("0x" . hex[0]))
  }
  return ObjDir
}

;返回以日期命名的路径“C:\xxx\xxx\MMddHHmmss”没有后缀
GetNewFilePath()
{
  NewDirName := GetObjDir()
  NewDirName .= "/"
  NewDirName .= FormatTime(, "MMddHHmmss")
  return NewDirName
}

;后台执行单条CMD命令并取得返回值
RunWaitOne(command) {
  shell := ComObject("WScript.Shell")
  ; 通过 cmd.exe 执行单条命令
  exec := shell.Exec(A_ComSpec " /C " command)
  ; 读取并返回命令的输出
  return exec.StdOut.ReadAll()
}

;后台执行多条CMD命令并取得返回值
RunWaitMany(commands) {
  shell := ComObject("WScript.Shell")
  ; 打开 cmd.exe 禁用命令回显
  exec := shell.Exec(A_ComSpec " /Q /K echo off")
  ; 发送并执行命令, 使用新行分隔
  exec.StdIn.WriteLine(commands "`nexit")  ; 总是在最后退出!
  ; 读取并返回所有命令的输出
  return exec.StdOut.ReadAll()
}

;在windows托盘显示信息
;传入参数为：
;title-string 信息标题
;infoMsg-string 信息内容 
;持续时间 int(负数)
SetWindowsInfo(title,infoMsg,time){
    Persistent
    TrayTip(infoMsg, title)
    SetTimer(TrayTip,time)
}

SetWindowsWarning(title,infoMsg,time){
    Persistent
    TrayTip(infoMsg, title,2)
    SetTimer(TrayTip,time)
}

SetWindowsError(title,infoMsg){
    Persistent
    TrayTip(infoMsg, title,3)
    ;SetTimer(TrayTip,time)
}


; 虚拟桌面切换函数,win10与win11所用的DLL不同. win10中Name桌面的name获取异常，已经注释
; AutoHotkey v2 script
SetWorkingDir(A_ScriptDir)

; Path to the DLL, relative to the script
VDA_PATH := A_ScriptDir . "\VirtualDesktopAccessor.dll"
hVirtualDesktopAccessor := DllCall("LoadLibrary", "Str", VDA_PATH, "Ptr")

GetDesktopCountProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "GetDesktopCount", "Ptr")
GoToDesktopNumberProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "GoToDesktopNumber", "Ptr")
GetCurrentDesktopNumberProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "GetCurrentDesktopNumber", "Ptr")
IsWindowOnCurrentVirtualDesktopProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "IsWindowOnCurrentVirtualDesktop", "Ptr")
IsWindowOnDesktopNumberProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "IsWindowOnDesktopNumber", "Ptr")
MoveWindowToDesktopNumberProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "MoveWindowToDesktopNumber", "Ptr")
IsPinnedWindowProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "IsPinnedWindow", "Ptr")
GetDesktopNameProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "GetDesktopName", "Ptr")
SetDesktopNameProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "SetDesktopName", "Ptr")
CreateDesktopProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "CreateDesktop", "Ptr")
RemoveDesktopProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "RemoveDesktop", "Ptr")

; On change listeners
RegisterPostMessageHookProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "RegisterPostMessageHook", "Ptr")
UnregisterPostMessageHookProc := DllCall("GetProcAddress", "Ptr", hVirtualDesktopAccessor, "AStr", "UnregisterPostMessageHook", "Ptr")

GetDesktopCount() {
    global GetDesktopCountProc
    count := DllCall(GetDesktopCountProc, "Int")
    return count
}

MoveCurrentWindowToDesktop(number) {
    global MoveWindowToDesktopNumberProc, GoToDesktopNumberProc
    activeHwnd := WinGetID("A")
    DllCall(MoveWindowToDesktopNumberProc, "Ptr", activeHwnd, "Int", number, "Int")
    DllCall(GoToDesktopNumberProc, "Int", number, "Int")
}

GoToPrevDesktop() {
    global GetCurrentDesktopNumberProc, GoToDesktopNumberProc
    current := DllCall(GetCurrentDesktopNumberProc, "Int")
    last_desktop := GetDesktopCount() - 1
    ; If current desktop is 0, go to last desktop
    if (current = 0) {
        MoveOrGotoDesktopNumber(last_desktop)
    } else {
        MoveOrGotoDesktopNumber(current - 1)
    }
    return
}

GoToNextDesktop() {
    global GetCurrentDesktopNumberProc, GoToDesktopNumberProc
    current := DllCall(GetCurrentDesktopNumberProc, "Int")
    last_desktop := GetDesktopCount() - 1
    ; If current desktop is last, go to first desktop
    if (current = last_desktop) {
        MoveOrGotoDesktopNumber(0)
    } else {
        MoveOrGotoDesktopNumber(current + 1)
    }
    return
}

GoToDesktopNumber(num) {
    global GoToDesktopNumberProc
    DllCall(GoToDesktopNumberProc, "Int", num, "Int")
    return
}
MoveOrGotoDesktopNumber(num) {
    ; If user is holding down Mouse left button, move the current window also
    if (GetKeyState("LButton")) {
        MoveCurrentWindowToDesktop(num)
    } else {
        GoToDesktopNumber(num)
    }
    return
}
GetDesktopName(num) {
    global GetDesktopNameProc
    utf8_buffer := Buffer(1024, 0)
    ran := DllCall(GetDesktopNameProc, "Int", num, "Ptr", utf8_buffer, "Ptr", utf8_buffer.Size, "Int")
    name := StrGet(utf8_buffer, 1024, "UTF-8")
    return name
}
SetDesktopName(num, name) {
    global SetDesktopNameProc
    OutputDebug(name)
    name_utf8 := Buffer(1024, 0)
    StrPut(name, name_utf8, "UTF-8")
    ran := DllCall(SetDesktopNameProc, "Int", num, "Ptr", name_utf8, "Int")
    return ran
}
CreateDesktop() {
    global CreateDesktopProc
    ran := DllCall(CreateDesktopProc, "Int")
    return ran
}
RemoveDesktop(remove_desktop_number, fallback_desktop_number) {
    global RemoveDesktopProc
    ran := DllCall(RemoveDesktopProc, "Int", remove_desktop_number, "Int", fallback_desktop_number, "Int")
    return ran
}

; SetDesktopName(0, "It works! 🐱")

DllCall(RegisterPostMessageHookProc, "Ptr", A_ScriptHwnd, "Int", 0x1400 + 30, "Int")
OnMessage(0x1400 + 30, OnChangeDesktop)

;win10中Name获取异常，win11可以更换DLL之后尝试
OnChangeDesktop(wParam, lParam, msg, hwnd) {
    Critical(1)
    OldDesktop := wParam + 1
    NewDesktop := lParam + 1
    ;Name := GetDesktopName(NewDesktop - 1)

    ; Use Dbgview.exe to checkout the output debug logs
    ;OutputDebug("Desktop changed to " Name " from " OldDesktop " to " NewDesktop)
    OutputDebug("Desktop changed from " OldDesktop " to " NewDesktop)
    ; TraySetIcon(".\Icons\icon" NewDesktop ".ico")
}

;=====================================================================o
;                         系统设置                                     |
;=====================================================================o

;切换虚拟桌面

#1::GotoDesktopNumber(0)
#2::GotoDesktopNumber(1)
#3::GotoDesktopNumber(2)
#4::GotoDesktopNumber(3)
#5::GotoDesktopNumber(4)
#6::GotoDesktopNumber(5)
#7::GotoDesktopNumber(6)
#8::GotoDesktopNumber(7)
#9::GotoDesktopNumber(8)



;新建空白markdown文档
^+m::
{
    filePath := GetNewFilePath()
    filePath .= ".md"
    FileAppend("", filePath)
}
;新建空白txt文档
^+i::
{
    filePath := GetNewFilePath()
    filePath .= ".txt"
    FileAppend("", filePath)
}
;新建空白无后缀文件
^+u::
{
    filePath := GetNewFilePath()
    FileAppend("", filePath)
}

;关闭显示器,不要多次连续触发此快捷键，唤醒屏幕需要等一会
#+l::
{
  static isMonitorOff := false
  if isMonitorOff
    {
      ;使用monitor on命令时需要bios支持，所以使用sendkey命令唤醒兼容性更好
      exitCode := RunWait("nircmd.exe sendkeypress ctrl")
      if !exitCode 
        ;异或操作实现切换bool值
        isMonitorOff ^= true
    }
  else
    {
      exitCode := RunWait("nircmd.exe monitor off")
      if !exitCode 
        isMonitorOff ^= true
    }
  
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
#+q::
{
	exitCode := RunWait("C:\Windows\System32\powercfg.exe setactive 6ed08f3e-52a1-4d96-9ae0-2f6619c8cdfd", ,"Hide")
  if !exitCode
    SetWindowsInfo("电源计划切换","切换至节电模式",-3000)
  else
    SetWindowsError("电源计划切换","切换节电模式失败")
}

;切换到平衡模式
#+w::
{
	exitCode := RunWait("C:\Windows\System32\powercfg.exe setactive 381b4222-f694-41f0-9685-ff5bb260df2e", ,"Hide")
  if !exitCode 
    SetWindowsInfo("电源计划切换","切换至平衡模式",-3000)
  else
    SetWindowsError("电源计划切换","切换平衡模式失败")
}

;切换到高性能模式
#+e::
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

