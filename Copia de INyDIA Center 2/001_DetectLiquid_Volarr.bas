'#Reference {B1A8DBB6-5B85-4DCA-818C-72193240147D}#1.0#0#XUtool.exe#XUtool Object Library
'#Reference {205EF3CF-98A6-4B54-974E-ADB070AE7EA1}#1.0#0#C:\Lirix\lib\xcfg.dll#Xiril: Create Driver Cfg
'#Reference {49ACF446-FF35-432F-8D85-8C427A295515}#1.9#0#C:\Lirix\lib\cuf.dll#Common Utility Functions
'#Reference {1F628F7C-2946-4024-981B-AC9E15BD0C9E}#3.7#0#C:\Lirix\lib\xsl.dll#Xiril Standard Functions
'#Reference {B1A8DBB6-5B85-4DCA-818C-72193240147D}#1.0#0#C:\Lirix\xtool.exe#XTool Object Lib
'#Reference {2ADE84F3-A0C1-451D-8C4E-EFB798F13EB3}#1.0#0#C:\Lirix\lib\xvc.dll#Xiril Vacuum Controller
'#Reference {39F316B7-2674-454F-84A9-BB3FCC5FB426}#3.0#0#C:\Lirix\lib\h96.dll#Xiril H96 Functions
'#Reference {567AF758-BD06-43A2-807B-08641C2F577A}#4.2#0#C:\Lirix\lib\xhl75.dll#Xiril75 HighLevel Functions
'#Reference {5A276AC0-0B95-42D6-88B3-2FA8D13CFF62}#3.2#0#C:\Lirix\lib\xhl100.dll#Xiril100 HighLevel Functions
'#Reference {4B139FE6-E343-466A-AC2C-86A2F32F8168}#1.0#0#C:\Lirix\lib\xwf.dll#Xiril Wedge Functions
'#Reference {1F628F7C-2946-4024-981B-AC9E15BD0C9E}#3.2#0#C:\Lirix\lib\xsl.dll#Xiril Standard Functions
'#Reference {5B7D8408-D456-4F7F-89DC-3E748EBE52F3}#2.0#0#C:\Lirix\lib\xpli.dll#Xiril Piplist Interpreter
'#Reference {5D385497-B597-48C7-AB19-D4F96BFAFE0A}#1.1#0#C:\Lirix\lib\xbaf.dll#Xiril Balance Functions
'#Reference {EA922DDD-D55C-4165-8CE8-9DEAB9BCB7AC}#1.0#0#C:\Lirix\lib\xrs232.dll#XCommRS232
'#Reference {49ACF446-FF35-432F-8D85-8C427A295515}#1.2#0#C:\Lirix\lib\cuf.dll#Common Utility Functions
Option Explicit
'#Uses "..\global\xbase.bas"


'Declaration of global variables

DIM NumberOfTips AS LONG
DIM ProcessSourceTubes AS LONG
DIM VolOfQuota AS LONG
DIM MinLastQuota AS LONG
DIM IncompleteQuota AS LONG

'#Uses "..\global\xbase.bas"
'#Uses "..\global\xbaseMain.bas"
'#Uses "..\global\xbase100.bas"
'#Uses "..\global\customlib.bas"
'#Uses "..\global\functionlib.bas"

Sub Init()
	Call InitGlobalObjVar
	Call InitGlobalObjVarX100
End Sub

Sub Start(ProcessLayoutName As String)
    Call xs.SetOutputSwitch(ResSwitch1,,, S_ON, "Switch blue diodes On")
    Call xh.DeleteRunList
    Call xs.InitSystem
    'create rack placement for current process
    xs.CreateProcessLayout (ProcessLayoutName)
End Sub

Sub Cleanup()
    Dim State As Long
    Call xs.SetOutputSwitch(ResSwitch1,,, S_OFF, "Switch blue diodes Off")
   'store communication runlog filename in interface file
    Call ff.WriteINI(AppPath & ".\data\app_in.dat", "General", "RunLog", xh.GetCurrentLogFileName)
    'set state when macro is finished
    If xs.AbortRunDetected = True Then
        State = 2       'run aborted
    Else
        State = 0       'run finished
    End If
    Call ff.WriteINI(AppPath & ".\data\app_in.dat", "General", "State", CStr(State)) 'finish or aborted
    Call xh.StoreLogFile

	Call ClearObj
	Call ClearObjX100
End Sub



'Add the global routine code here


Sub DetectLiquid_VolArr_P00_A0000
	Dim tmp as String
	tmp = "DetectLiquid_VolArr:"  & vbCrLf & vbCrLf
	tmp = tmp & "NumberOfTips:    " & Str(TPLNumberOfTipsTPL) & vbCrLf
	tmp = tmp & "ProcessSourceTubes:    " & Str(TPLProcessSourceTubesTPL) & vbCrLf
	tmp = tmp & "VolOfQuota:    " & Str(TPLVolOfQuotaTPL) & vbCrLf
	tmp = tmp & "MinLastQuota:    " & Str(TPLMinLastQuotaTPL) & vbCrLf
	tmp = tmp & "IncompleteQuota:    " & Str(TPLIncompleteQuotaTPL) & vbCrLf
	MsgBox tmp
End Sub



Sub InitVariables

	NumberOfTips = 4
	ProcessSourceTubes = 24
	VolOfQuota = 500
	MinLastQuota = 100
	IncompleteQuota = 1

End Sub

Sub InputVariables
INPUTERROR:
	Begin Dialog UserDialog 340, 130, "Variables"
		Text      10,  12, 120,  20, "ProcessSourceTubes:"
		TextBox  140,  10, 180,  20, .Text1$
		Text      10,  32, 120,  20, "VolOfQuota:"
		TextBox  140,  30, 180,  20, .Text2$
		Text      10,  52, 120,  20, "MinLastQuota:"
		TextBox  140,  50, 180,  20, .Text3$
		Text      10,  72, 120,  20, "IncompleteQuota:"
		TextBox  140,  70, 180,  20, .Text4$
		OKButton 220, 100, 100,  20
	End Dialog

	Dim Dlg As UserDialog

	dlg.Text1$ = Str$(ProcessSourceTubes)
	dlg.Text2$ = Str$(VolOfQuota)
	dlg.Text3$ = Str$(MinLastQuota)
	dlg.Text4$ = Str$(IncompleteQuota)

	Dialog dlg

	ProcessSourceTubes = CLng(dlg.Text1$)
	If ProcessSourceTubes < 4 OR ProcessSourceTubes > 96 Then
		MsgBox "Invalid Value 'ProcessSourceTubes', Allowed Range: [4,96]"
		GoTo INPUTERROR
	End If
	VolOfQuota = CLng(dlg.Text2$)
	If VolOfQuota < 100 OR VolOfQuota > 1000 Then
		MsgBox "Invalid Value 'VolOfQuota', Allowed Range: [100,1000]"
		GoTo INPUTERROR
	End If
	MinLastQuota = CLng(dlg.Text3$)
	If MinLastQuota < 10 OR MinLastQuota > 1000 Then
		MsgBox "Invalid Value 'MinLastQuota', Allowed Range: [10,1000]"
		GoTo INPUTERROR
	End If
	IncompleteQuota = CLng(dlg.Text4$)
	If IncompleteQuota < 0 OR IncompleteQuota > 1 Then
		MsgBox "Invalid Value 'IncompleteQuota', Allowed Range: [0,1]"
		GoTo INPUTERROR
	End If
End Sub

Sub Main

	'Creating the main object for work
	CreateXToolObject ""

	'Initialisation code
	Init

	'Variable Initialisation code
	InitVariables

	'Variable Input code
	InputVariables

	'Startup code
	Start ""

	'Routines
	DetectLiquid_VolArr_P00_A0000

	'Cleanup code
	Cleanup

End Sub
