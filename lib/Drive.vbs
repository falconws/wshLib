Option Explicit

'�{�����[���������Ƀh���C�u���^�[���擾����
'@param volumeName
'	�擾�Ώۂ̃h���C�u�̃{�����[����
'@return
'	�h���C�u���^�[ (C: ��)
Function getDriveLetterByVolumeName(volumeName)

	Dim objFso
	Dim drive

	Set objFso = CreateObject("Scripting.FileSystemObject")

	For Each drive In objFso.Drives
		If drive.IsReady Then
			If drive.VolumeName = volumeName Then
				getDriveLetterByVolumeName = drive.Path
			End If
		End If
	Next
	
	Set objFso = Nothing
End Function