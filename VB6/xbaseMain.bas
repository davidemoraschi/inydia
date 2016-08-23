Attribute VB_Name = "xbaseMain"
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
'*************************************************************************
' Xiril Global Module
'*************************************************************************
Option Explicit

'#Uses ".\xbase.bas"

Public xs As XirilStandardFunctions
Public pl As XirilPipListInterpreter
Public rs As XCommRS232.XRS232
Public aa As cuf.AllAround
Public ff As cuf.FileFunctions
Public uf As cuf.UtilityFunctions
Public AppPath$

Public Enum eYesNo
        YES = 1
        NO = 0
End Enum

Public Enum eDiTiType
    TIPS_250ul = &H1
    TIPS_1000ul = &H2
    TIPS_20ul = &H80
End Enum

Public Enum eDiTiTypeNew
    Rainin_250ul = &H1
    Rainin_1000ul = &H2
    Rainin_20ul = &H80
    Rainin_250ulF = &H100
    Rainin_1000ulF = &H200
End Enum

Public Enum eChannels
    AllChannels = -1
    Channel1 = &H1      '1
    Channel2 = &H2      '2
    Channel3 = &H4      '4
    Channel4 = &H8      '8
    Channel5 = &H10     '16
    Channel6 = &H20     '32
    Channel7 = &H40     '64
    Channel8 = &H80     '128
End Enum

Public Enum eShortChannels
    C1 = &H1    '1
    C2 = &H2    '2
    C3 = &H4    '4
    C4 = &H8    '8
    C5 = &H10   '16
    C6 = &H20   '32
    C7 = &H40   '64
    C8 = &H80   '128
End Enum

Public Enum eCSVRackInfoOptions
        ActualDateTime = 1
        RackName = 2
        RackID = 3
        RackPos = 4
        RackPosID = 5
        LqdInVol = 6
        NoOfLiquidErrFlags = 7
        EmptyField = 8
        unused = 0
End Enum

Public Enum eFileOptions
        DeleteExistingFileFirst = 0
        AppendToExistingFile = 1
End Enum

Public Enum ePosInfoOptions
        Source = 0
        Destination = 1
End Enum

Public Enum eAspDispModeOptions
        ZPos = 0
        Track = 2
        FromLastLqdLevel = 1
        ClotDetection = 4
        ClotDetect_Ignore = 20
        ClotDetect_Retry = 12
        ClotErr_Auto_Ignore = &H10      '16
        ClotErr_Auto_Retry = &H8        '8
End Enum

Public Enum eDispModeOptions
        RackZDisp = 0
        DispZPos = 1
        AtLastLqdLevel = 4
End Enum

Public Enum eBarcodeErrorOptions
        NoCheck = 0
        CheckMissingBC = 1
        CheckDuplicateBC = 2
End Enum

Public Enum eCreateSampleID
        CS_NO = 0
        CS_RUNNING_NUMBER = 1
        CS_RACKID_RACKPOS = 2
End Enum


Public Sub InitGlobalObjVar()
    Set xs = New XirilStandardFunctions
    Set pl = New XirilPipListInterpreter
        Set rs = New XCommRS232.XRS232
        Set aa = New cuf.AllAround
        Set ff = New cuf.FileFunctions
        Set uf = New cuf.UtilityFunctions

        'read AppPath from Registry
        AppPath = aa.RegValueGet(HKEY_LOCAL_MACHINE, "Software\Xiril\Lirix\Install", "AppPath")

    Call xs.InitLib(XToolApp, Robot, False)
    Call pl.InitLib(XToolApp, Robot, xs)
End Sub

Public Sub ClearObj()
        Set aa = Nothing
        Set ff = Nothing
        Set uf = Nothing

    Set rs = Nothing
    Set pl = Nothing
    Set xs = Nothing
End Sub
