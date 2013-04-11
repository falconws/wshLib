Option Explicit

'pathで与えられたフォルダを作成する。親フォルダが
'存在しない場合自動的に作成する
'@param objFso
'	FileSystemObject
'@param strFolder
'	作成するフォルダのパス
Sub makeFolder(objFso, strFolder)

	Dim strParent
	strParent = objFso.GetParentFolderName(strFolder)
	If Not objFso.FolderExists(strParent) Then
		makeFolder objFso, strParent
	End If
	If Not objFso.FolderExists(strFolder) Then
		objFso.CreateFolder(strFolder)
	End If
End Sub