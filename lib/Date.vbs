Option Explicit

'YYYY_MM_DD 形式の文字列を取得する
'@param 日付を表す文字列
'@return 引数をYYYY_MM_DDに変換した文字列
Function getYYYY_MM_DD(date)
	getYYYY_MM_DD = Year(date) & "_" & Month(date) & "_" & Day(date)
End Function

