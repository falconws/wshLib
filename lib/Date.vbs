Option Explicit

'YYYY_MM_DD �`���̕�������擾����
'@param ���t��\��������
'@return ������YYYY_MM_DD�ɕϊ�����������
Function getYYYY_MM_DD(date)
	getYYYY_MM_DD = Year(date) & "_" & Month(date) & "_" & Day(date)
End Function

