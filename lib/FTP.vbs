Option Explicit

'�����Ŏ󂯎����FTP�R�}���h�t�@�C�������s����
'@param fileName �������s����FTP�R�}���h�t�@�C���t���p�X
Sub autoFTP(fileName)
	Dim objShell
	
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run "ftp.exe -s:" & fileName, 0, True
	Set objShell = Nothing
End Sub