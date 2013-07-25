Option Explicit

'Windows OS�̃o�[�W������Ԃ�
'���݂��Ȃ��ꍇ�����I�ɍ쐬����
'@return
'	OS�̃o�[�W����������������
Function getOS()
	Dim OSInfoCollection, OSInfo, retOS

	Set OSInfoCollection = GetObject("winmgmts:").InstancesOf("Win32_OperatingSystem")

	For Each OSInfo In OSInfoCollection
    	retOS = OSInfo.Caption
	Next
	
	Set OSInfoCollection = Nothing
	getOS = retOS
End Function

'AllUsers�̃X�^�[�g�A�b�v�փv���O������o�^����
'@param workDir ���s����v���O�������i�[����Ă���t�H���_�ւ̐�΃p�X
'@param execFileName ���s����v���O�����̃t�@�C����
'@param windowStyle �V���[�g�J�b�g���s���̃E�B���h�E�T�C�Y 1:���� 3:�ő剻 7:�ŏ���
Sub registAllUsersStartup(workDir, execFileName, windowStyle)
	Dim objWsh
	Dim startupDir, oShellLink
	
	Set objWsh = CreateObject("WScript.Shell")
	
	rem �X�^�[�g�A�b�v�f�B���N�g���̎擾
	startupDir = objWsh.SpecialFolders("AllUsersStartup")

	rem �X�^�[�g�A�b�v�֓o�^
	Set oShellLink = objWsh.CreateShortcut(startupDir & "\" & execFileName & ".lnk")
	oShellLink.TargetPath = workDir & "\" & execFileName
	oShellLink.WorkingDirectory = workDir

	oshellLink.WindowStyle = windowStyle

	oshellLink.Save
	
	Set objWsh = Nothing
	Set oShellLink = Nothing
End Sub

'�w�肳�ꂽ�v���Z�X���̃v���Z�X�������I������
'@param processName �v���Z�X��
Function killProcessByName(processName)
	Dim objLocator, objService, colProcSet, objProc
	
	'�I�u�W�F�N�g�̐���
	Set objLocator = WScript.CreateObject("WbemScripting.SWbemLocator")
	Set objService = objLocator.ConnectServer
	
	Set colProcSet = objService.ExecQuery("Select * From Win32_Process Where Caption='" & processName & "'")
	For Each objProc In colProcSet
		objProc.Terminate
	Next
End Function

'OS���V���b�g�_�E������
Sub shutdown()
	Dim objShutdown, objOS, objSystem
	
	' �V���b�g�_�E���I�u�W�F�N�g�擾
	Set objShutdown = GetObject("winmgmts:{impersonationLevel = impersonate, (Shutdown)}")
	' OS�̃I�u�W�F�N�g�擾
	Set objOS = objShutdown.InstancesOf("Win32_OperatingSystem")
 
	' �V���b�g�_�E��
	For Each objSystem In objOS
    	objSystem.Win32Shutdown 8
	Next
End Sub

'OS���ċN������
Sub reboot()
	Dim objShutdown, objOS, objSystem
	
	' �V���b�g�_�E���I�u�W�F�N�g�擾
	Set objShutdown = GetObject("winmgmts:{impersonationLevel = impersonate, (Shutdown)}")
	' OS�̃I�u�W�F�N�g�擾
	Set objOS = objShutdown.InstancesOf("Win32_OperatingSystem")
 
	' �ċN��
	For Each objSystem In objOS
    	objSystem.Win32Shutdown 2
	Next
End Sub