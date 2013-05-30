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


'grepコマンドもどき
'テキストファイルから特定の文字列を検索する
'参考: http://www.upken.jp/kb/vbscript_grep.html
'@param strSearch 検索文字列
'@param filename 検索対象のファイル名
'@param argument grepコマンドの-v(キーワード除外)か-i(大文字小文字無視)
Function grep(strSearch, filename, argument)
	Const ForReading = 1
	Dim objFso, myReg
	Dim myRead, myLine
	Dim isIgnoreCase, isIgnorePattern
	Set objFso = CreateObject("Scripting.FileSystemObject")

	isIgnoreCase = False
	isIgnorePattern = False
	
	If argument = "-i" Then
		isIgnoreCase = True
	ElseIf argument = "-v" Then
		isIgnorePattern = True
	End If

	Set myReg = new RegExp
	myReg.Pattern = strSearch
	myReg.Global = False
	myReg.IgnoreCase = isIgnoreCase

	If Not objFso.FileExists(filename) Then
		Err.Raise 513, "Util.grep()", "File Not Found." & filename
	End If
	
	Set myRead = objFso.OpenTextFile(filename, ForReading)
	Do While myRead.AtEndOfStream = False
		myLine = myRead.ReadLine
		If myReg.Test(myLine) Then
			If Not isIgnorePattern Then
				grep = grep & myLine & vbNewLine
			End If
		Else
			If isIgnorePattern Then
				grep = grep & myLine & vbNewLine
			End If
		End If
	Loop
	myRead.Close
	
	' 余分な末尾の改行コードを削除する
	' http://gallery.technet.microsoft.com/scriptcenter/d7ae5c5d-38be-4cfc-b415-82e5a79e870e
	If Right(grep, 2) = vbCrLf Then
		grep = Left(grep, Len(grep) - 2)
	End If
End Function