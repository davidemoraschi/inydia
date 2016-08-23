'#Reference {B1A8DBB6-5B85-4DCA-818C-72193240147D}#1.0#0#XUtool.exe#XUtool Object Library
Option Explicit
'#Uses "C:\Lirix\data\global\xbase.bas"

'NOTE: The main object is declared. If an error occurs
'please check if the reference to the XUtool library is set.
'(Menu Edit/References)
DIM TPLNumberOfTipsTPL AS LONG
DIM TPLProcessSourceTubesTPL AS LONG
DIM TPLVolOfQuotaTPL AS LONG
DIM TPLMinLastQuotaTPL AS LONG
DIM TPLIncompleteQuotaTPL AS LONG

'BEGINCODEONCE

'Add the global routine code here

'ENDCODE

'BEGINCODE

Sub DetectLiquid_VolArr
	Dim tmp as String
	tmp = "DetectLiquid_VolArr:"  & vbCrLf & vbCrLf
	tmp = tmp & "NumberOfTips:    " & Str(TPLNumberOfTipsTPL) & vbCrLf
	tmp = tmp & "ProcessSourceTubes:    " & Str(TPLProcessSourceTubesTPL) & vbCrLf
	tmp = tmp & "VolOfQuota:    " & Str(TPLVolOfQuotaTPL) & vbCrLf
	tmp = tmp & "MinLastQuota:    " & Str(TPLMinLastQuotaTPL) & vbCrLf
	tmp = tmp & "IncompleteQuota:    " & Str(TPLIncompleteQuotaTPL) & vbCrLf
	MsgBox tmp
End Sub

'ENDCODE

Sub Main

	CreateXToolObject ""
	TPLNumberOfTipsTPL = 4
	TPLProcessSourceTubesTPL = 24
	TPLVolOfQuotaTPL = 500
	TPLMinLastQuotaTPL = 100
	TPLIncompleteQuotaTPL = 1
	DetectLiquid_VolArr
End Sub
