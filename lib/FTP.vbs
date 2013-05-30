Option Explicit

'引数で受け取ったFTPコマンドファイルを実行する
'@param fileName 自動実行するFTPコマンドファイルフルパス
Sub autoFTP(fileName)
	Dim objShell
	
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run "ftp.exe -s:" & fileName, 0, True
	Set objShell = Nothing
End Sub