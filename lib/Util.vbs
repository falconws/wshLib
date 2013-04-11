Option Explicit

' robocopy.exeをインストールする
' @param robocopy
'        インストールするrobocopy.exeへの絶対パス
Sub installRobocopy(robocopy)
	Const DOUBLE_QUOTE = """"
	Const SYSTEM32 = "C:\WINDOWS\system32"
	
	Dim objWsh, objExec
	Dim pwd
	
	Set objWsh = CreateObject("WScript.Shell")
	' system32 ディレクトリは読み取り専用で、CopyFileメソッドが失敗する
	' OSのコピーコマンドを使って回避する。

	' UNC パスからのコピーをサポートする為一時的にカレントディレクトリ変更
	' カレントディレクトリ
	' pwd = objWsh.CurrentDirectory
	' objWsh.CurrentDirectory = "C:\"
	
	' ファイルコピー
	Set objExec = objWsh.Exec("cmd /c copy /Y " & DOUBLE_QUOTE & _
			robocopy & DOUBLE_QUOTE & " " & SYSTEM32)
	If Not objExec.StdErr.AtEndOfStream Then
		MsgBox objExec.StdErr.ReadAll()
	End If

	' カレントディレクトリを元に戻す
	' objWsh.CurrentDirectory = pwd
	
	Set objWsh = Nothing
	Set objExec = Nothing
End Sub