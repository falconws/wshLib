Option Explicit

'ボリューム名を元にドライブレターを取得する
'@param volumeName
'	取得対象のドライブのボリューム名
'@return
'	ドライブレター (C: 等)
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