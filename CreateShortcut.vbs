Set oWS = WScript.CreateObject("WScript.Shell")
sLinkFile = oWS.SpecialFolders("Desktop") & "\���AI�����°�.lnk"
Set oLink = oWS.CreateShortcut(sLinkFile)

' ��ȡ��ǰ�ű����ڵ�Ŀ¼
sCurrentDir = oWS.CurrentDirectory

oLink.TargetPath = sCurrentDir & "\���AI����.exe" ' Ŀ���ļ�·��
oLink.WindowStyle = 1
oLink.IconLocation = sCurrentDir & "\���AI����.exe, 0" ' ��ݷ�ʽͼ�꣬��������Դ�ͼ��
oLink.Description = "���AI�����°�" ' ��ݷ�ʽ����
oLink.WorkingDirectory = sCurrentDir ' Ŀ���ļ�����Ŀ¼
oLink.Save