Option Explicit

'path�ŗ^����ꂽ�t�H���_���쐬����B�e�t�H���_��
'���݂��Ȃ��ꍇ�����I�ɍ쐬����
'@param objFso
'	FileSystemObject
'@param strFolder
'	�쐬����t�H���_�̃p�X
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