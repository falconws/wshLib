Option Explicit

' robocopy.exe���C���X�g�[������
' @param robocopy
'        �C���X�g�[������robocopy.exe�ւ̐�΃p�X
Sub installRobocopy(robocopy)
	Const DOUBLE_QUOTE = """"
	Const SYSTEM32 = "C:\WINDOWS\system32"
	
	Dim objWsh, objExec
	Dim pwd
	
	Set objWsh = CreateObject("WScript.Shell")
	' system32 �f�B���N�g���͓ǂݎ���p�ŁACopyFile���\�b�h�����s����
	' OS�̃R�s�[�R�}���h���g���ĉ������B

	' UNC �p�X����̃R�s�[���T�|�[�g����׈ꎞ�I�ɃJ�����g�f�B���N�g���ύX
	' �J�����g�f�B���N�g��
	' pwd = objWsh.CurrentDirectory
	' objWsh.CurrentDirectory = "C:\"
	
	' �t�@�C���R�s�[
	Set objExec = objWsh.Exec("cmd /c copy /Y " & DOUBLE_QUOTE & _
			robocopy & DOUBLE_QUOTE & " " & SYSTEM32)
	If Not objExec.StdErr.AtEndOfStream Then
		MsgBox objExec.StdErr.ReadAll()
	End If

	' �J�����g�f�B���N�g�������ɖ߂�
	' objWsh.CurrentDirectory = pwd
	
	Set objWsh = Nothing
	Set objExec = Nothing
End Sub


'grep�R�}���h���ǂ�
'�e�L�X�g�t�@�C���������̕��������������
'�Q�l: http://www.upken.jp/kb/vbscript_grep.html
'@param strSearch ����������
'@param filename �����Ώۂ̃t�@�C����
'@param argument grep�R�}���h��-v(�L�[���[�h���O)��-i(�啶������������)
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
	
	' �]���Ȗ����̉��s�R�[�h���폜����
	' http://gallery.technet.microsoft.com/scriptcenter/d7ae5c5d-38be-4cfc-b415-82e5a79e870e
	If Right(grep, 2) = vbCrLf Then
		grep = Left(grep, Len(grep) - 2)
	End If
End Function