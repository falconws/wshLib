Option Explicit

'�����Ŏ󂯎����FTP�R�}���h�t�@�C�������s����
'@param fileName �������s����FTP�R�}���h�t�@�C���t���p�X
'@return FTP�R�}���h�W���o��
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