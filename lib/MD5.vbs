Option Explicit

' �t�@�C����MD5�n�b�V���l���擾����
'@param filename �t�@�C���̃t���p�X
Function md5(filename)
	Dim MSXML, EL
	Set MD5 = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
	MD5.ComputeHash_2(readBinaryFile(filename))

	Set MSXML = CreateObject("MSXML2.DOMDocument")
	Set EL = MSXML.CreateElement("tmp")
	EL.DataType = "bin.hex"
	EL.NodeTypedValue = MD5.Hash
	md5 = EL.Text
End Function

' �o�C�i���t�@�C����ǂݍ���
'@param filename �t�@�C���̃t���p�X
Function readBinaryFile(filename)
	Const adTypeBinary = 1
	Dim objStream
	Set objStream = CreateObject("ADODB.Stream")
	objStream.Type = adTypeBinary
	objStream.Open
	objStream.LoadFromFile filename
	readBinaryFile = objStream.Read
	objStream.Close
	Set objStream = Nothing
End Function