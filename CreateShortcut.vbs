Set oWS = WScript.CreateObject("WScript.Shell")
sLinkFile = oWS.SpecialFolders("Desktop") & "\昊天AI换脸新版.lnk"
Set oLink = oWS.CreateShortcut(sLinkFile)

' 获取当前脚本所在的目录
sCurrentDir = oWS.CurrentDirectory

oLink.TargetPath = sCurrentDir & "\昊天AI换脸.exe" ' 目标文件路径
oLink.WindowStyle = 1
oLink.IconLocation = sCurrentDir & "\昊天AI换脸.exe, 0" ' 快捷方式图标，假设程序自带图标
oLink.Description = "昊天AI换脸新版" ' 快捷方式描述
oLink.WorkingDirectory = sCurrentDir ' 目标文件所在目录
oLink.Save