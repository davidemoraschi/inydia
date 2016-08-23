'#Reference {39F316B7-2674-454F-84A9-BB3FCC5FB426}#3.0#0#C:\Lirix\lib\h96.dll#Xiril H96 Functions
'#Reference {5D385497-B597-48C7-AB19-D4F96BFAFE0A}#1.1#0#C:\Lirix\lib\xbaf.dll#Xiril Balance Functions
'#Reference {EA922DDD-D55C-4165-8CE8-9DEAB9BCB7AC}#1.0#0#C:\Lirix\lib\xrs232.dll#XCommRS232
'#Reference {49ACF446-FF35-432F-8D85-8C427A295515}#1.9#0#C:\Lirix\lib\cuf.dll#Common Utility Functions
'#Reference {9F61CC71-57A4-45C6-8BA6-6BE7D5A504CE}#1.0#0#C:\Lirix\XLogin.exe#Common Login Module
'#Reference {567AF758-BD06-43A2-807B-08641C2F577A}#4.3#0#C:\Lirix\lib\xhl75.dll#Xiril75 HighLevel Functions
'#Reference {5A276AC0-0B95-42D6-88B3-2FA8D13CFF62}#3.3#0#C:\Lirix\lib\xhl100.dll#Xiril100 HighLevel Functions
'#Reference {205EF3CF-98A6-4B54-974E-ADB070AE7EA1}#1.0#0#C:\Lirix\lib\xcfg.dll#Xiril: Create Driver Cfg
'#Reference {4B139FE6-E343-466A-AC2C-86A2F32F8168}#1.0#0#C:\Lirix\lib\xwf.dll#Xiril Wedge Functions
'#Reference {2ADE84F3-A0C1-451D-8C4E-EFB798F13EB3}#1.1#0#C:\Lirix\lib\xvc.dll#Xiril Vacuum Controller
'#Reference {3751DA88-8B5A-401D-A237-551838981C26}#1.1#0#C:\Lirix\lib\xuf.dll#Xiril Utility Functions
'#Reference {1F628F7C-2946-4024-981B-AC9E15BD0C9E}#3.7#0#C:\Lirix\lib\xsl.dll#Xiril Standard Functions
'#Reference {5B7D8408-D456-4F7F-89DC-3E748EBE52F3}#2.0#0#C:\Lirix\lib\xpli.dll#Xiril Piplist Interpreter

Option Explicit
'#Uses ".\xbase.bas"
'#Uses ".\xbaseMain.bas"

'utility function library
Private m_oLoginObj As Xlogin.XirilLoginObject

'returns the name of the user which is currently logged in
Public Function Util_GetLoginName As String
	If m_oLoginObj Is Nothing Then
		Set m_oLoginObj = GetObject(,"XLogin.XirilLoginObject")
	End If
	If Not m_oLoginObj Is Nothing Then
		Util_GetLoginName = m_oLoginObj.Username
	Else
		Util_GetLoginName = "Unknown"
	End If
End Function
