Option Explicit

'path�ŗ^����ꂽ�t�H���_���쐬����B�e�t�H���_��
'���݂��Ȃ��ꍇ�����I�ɍ쐬����
'@param strFolder
'	�쐬����t�H���_�̃p�X
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