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