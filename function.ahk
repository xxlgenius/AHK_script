;实现自定义函数和维护全局变量

;=====================================================================o
;                         全局变量                                     |
;=====================================================================o

;windows的临时文件路径，如果跨平台需要写个环境判断
tempFile := Format("{1}\AHKTMP",A_AppData)

#SingleInstance Force
global   DEFAULT_IMAGE_FOLDER := "D:\Data\Programs\UserData\AutoHotkey\tempImages"
global  DEFAULT_LOG_FOLDER := "D:\Data\Programs\UserData\AutoHotkey\tempImages"
global local_img
global clip_type
;#Include, Gdip_All.ahk
; 压缩剪贴版图片

OnClipboardChange("ClipChanged")

# Persistent
ClipChanged(Type) {
clip_type := Type
}

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