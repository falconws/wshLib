Option Explicit

'引数で受け取ったFTPコマンドファイルを実行する
'@param fileName 自動実行するFTPコマンドファイルフルパス
'@return FTPコマンド標準出力
Function autoFTP(fileName)
	Dim objShell, objExec
	Dim strcmd
	
	Set objShell = CreateObject("WScript.Shell")
	strcmd = "ftp.exe -s:" & fileName
	Set objExec = objShell.Exec(strcmd)
	
	Do Until objExec.Status = 0
		WScript.Sleep 100
	Loop
	
	autoFTP = objExec.StdOut.ReadAll()
	
	Set objShell = Nothing
	Set objExec = Nothing
End Function