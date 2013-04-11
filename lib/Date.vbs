Option Explicit

'YYYY_MM_DD Œ`®‚Ì•¶š—ñ‚ğæ“¾‚·‚é
'@param “ú•t‚ğ•\‚·•¶š—ñ
'@return ˆø”‚ğYYYY_MM_DD‚É•ÏŠ·‚µ‚½•¶š—ñ
Function getYYYY_MM_DD(date)
	getYYYY_MM_DD = Year(date) & "_" & Month(date) & "_" & Day(date)
End Function

