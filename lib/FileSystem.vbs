Option Explicit

'pathで与えられたフォルダを作成する。親フォルダが
'存在しない場合自動的に作成する
'@param strFolder
'	作成するフォルダのパス
Sub makeFolder(strFolder)
	Dim objFso
	Dim strParent
	Set objFso = CreateObject("Scripting.FileSystemObject")
	strParent = objFso.GetParentFolderName(strFolder)
	If Not objFso.FolderExists(strParent) Then
		makeFolder objFso, strParent
	End If
	If Not objFso.FolderExists(strFolder) Then
		objFso.CreateFolder(strFolder)
	End If
End Sub