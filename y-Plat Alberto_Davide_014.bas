'#Reference {00025E01-0000-0000-C000-000000000046}#5.0#0#C:\WINDOWS\system32\dao360.dll#Microsoft DAO 3.6 Object Library
'#Reference {49ACF446-FF35-432F-8D85-8C427A295515}#1.9#0#C:\Lirix\lib\cuf.dll#Common Utility Functions
'#Reference {1F628F7C-2946-4024-981B-AC9E15BD0C9E}#3.7#0#C:\Lirix\lib\xsl.dll#Xiril Standard Functions
'#Reference {D63E0CE2-A0A2-11D0-9C02-00C04FC99C8E}#2.0#0#C:\WINDOWS\system32\msxml.dll#Microsoft XML, version 2.0
'*************************************************************************
'
' Xiril100 Basic Macro Template (basic without additional device support)
'
'*************************************************************************

Option Explicit

'#Uses "..\global\xbase.bas"
'#Uses "..\global\xbasemain.bas"
'#Uses "..\global\xbase100.bas"
'#Uses "..\global\customlib.bas"
'#Uses "..\global\functionlib.bas"
'#Uses "..\global\INyDIA_SourceTube.cls"
'#Uses "..\global\INyDIA_LogFile.cls"
'#Uses "..\macros\INyDIA_SourceTube.bas"


'DEFINE PUBLIC PROCESS VARIABLES HERE  -----------------------------------------------------------
Public ProcessSourceTubes As Long
Public VolOfQuota As Long

Const PosInfoFile = "C:\Lirix\data\piplist\posinfo.dat" '+ Format(Now,"yyyymmdd-hhnnss")+".dat"
Const MAX_RACK_POS = 96

'-------------------------------------------------------------------------------------------------

Sub Init()
	If SUBPROCESS_ACTIVE Then
		Exit Sub
	End If
	Call InitGlobalObjVar
	Call InitGlobalObjVarX100
End Sub

Sub Start(ProcessLayoutName As String)
	If SUBPROCESS_ACTIVE Then
		Exit Sub
	End If
    Call xs.SetOutputSwitch(ResSwitch1,,, S_ON, "Switch blue diodes On")
    Call xh.DeleteRunList
    Call xs.InitSystem
    'create rack placement for current process/macro
    xs.CreateProcessLayout (ProcessLayoutName)
End Sub

Sub Cleanup()
	If SUBPROCESS_ACTIVE Then
		Exit Sub
	End If
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
   ' Call xh.StoreLogFile

	Call ClearObj
	Call ClearObjX100
End Sub


'-------------------------------------------------------------------------------------------------
'Globla sub routines called from MAIN function
Function InitGlobalVariables() As Boolean
    'Set default values
    ProcessSourceTubes  = 12
    VolOfQuota = 500

INPUTERROR:
	Begin Dialog UserDialog 430,154,"Insert Process Varibales" ' %GRID:10,7,1,1
		OKButton 310,119,100,21
		CancelButton 190,119,100,21
		Text 10,35,140,21,"No of Source tubes"
		Text 10,63,100,28,"Volumen Al�quota"
		TextBox 180,35,60,21,.ProcessSourceTubes
		TextBox 110,70,60,21,.VolOfQuota
		Text 60,7,300,21,"Inicio",.header,2
		Text 10,98,60,21,"NOTA:",.Text1
		Text 80,98,350,14,"Volumen en microlitros",.Text2
	End Dialog

    Dim Dlg As UserDialog

    On Error GoTo CANCELPRESSED

    Dlg.ProcessSourceTubes$ = Str$(ProcessSourceTubes )
    Dlg.VolOfQuota$ = Str$(VolOfQuota)

    '**************************************************************************
    '* davide 5/11/2010: deshabilita la dialog box de par�metros al principio *
    '*------------------------------------------------------------------------*
    '**************************************************************************
    'Dialog Dlg
    '**************************************************************************
    '* davide 5/11/2010: deshabilita la dialog box de par�metros al principio *
    '*------------------------------------------------------------------------*
    '**************************************************************************

    ProcessSourceTubes  = CLng(Dlg.ProcessSourceTubes$)
    'If ProcessSourceTubes  < 0 Or NoOfCycles > 10 Then
     '   MsgBox "Invalid Value for 'NoOfCycles', Allowed Range: [0,10]"
      '  GoTo INPUTERROR
    'End If
    VolOfQuota = CLng(Dlg.VolOfQuota$)

    InitGlobalVariables = True
    Exit Function

CANCELPRESSED:
    InitGlobalVariables = False

End Function
'-------------------------------------------------------------------------------------------------



'-------------------------------------------------------------------------------------------------
'MAIN FUNCTION
'-------------------------------------------------------------------------------------------------
Sub main()
    Dim ProcessLayout$
'MacroSetting Start ++++++++++++++++++++++++++++++++++++++++++++++++++++
ProcessLayout = "TEST"
'MacroSetting End ++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Call CreateXToolObject(ProcessLayout)

    Const NumberOfTips As Integer = 8
    Const ProcessSourceTubes As Integer  = 24
	Const VolOfQuota As Long  = 500
    Const IncompleteQuota As Boolean  = True 	' ***** Propiedad de la clase INyDIA_SourceTube
    Const MinLastQuota As Long = 100 			' ***** Propiedad de la clase INyDIA_SourceTube

	Const fldLirix = "C:\Lirix"
	Const fchMuestras = fldLirix & "\data\muestras.dat"
	Const fchplateBC = fldLirix & "\data\plateBC.dat"
	Const fchbcdata = fldLirix & "\data\bcdata.dat"

	Static MAXDispensedTarget As Long
	Static MAXDispensedRack As Integer
	Static OverAllDispensedQuotas As Long

	Dim Dbs As New DAO.DBEngine, db As DAO.Database
	Dim InitVar As Boolean
	Dim CurrentTubeSet As SourceTubeSet, xvol As SamVolArr
	Dim DestRackName(3) As String, WellMatrix(96) As String, BC() As String
	Dim a As Integer, CurrentPos As Integer, CurrentRack As Integer
	Dim i As Long, LineCounter As Long, n As Long, vExecution_ID As Long, CurrentQuota As Long
	Dim PosInfoLine As String, strFilePath As String, StrXMLFileName As String, BCLine As String, BCNumber As String, BCRack As String, BCPos As Integer

	Dim SrcAssName As Long
	Dim vNumMuestras As Integer' vExecution_ID As Long, ,
	Dim aNumPlacas(1) As Long


    InitVar = InitGlobalVariables

    If InitVar = True Then
        Call Init
        Call Start(ProcessLayout)

    '***********************************************************************************
    '* davide 5/11/2010: proceso de lectura c�digo de barras de la(s) placa(s) destino *
    '*---------------------------------------------------------------------------------*
    '* genera un fichero muestras.dat con el n�mero de muestras, un plateBC.dat con el *
    '* c�digo de barras de 1 o 2 placas de destino                                     *
    '***********************************************************************************
	'START CODING HERE -----------------------------------------------------------------
	'Variable declaration
	'SUBPROCESS_ACTIVE = True
	'XToolApp.CallSub("C:\Lirix\data\process\Read_Tubes.bas","main")
    '***********************************************************************************
    '* davide 5/11/2010: proceso de lectura c�digo de barras de la(s) placa(s) destino *
    '*---------------------------------------------------------------------------------*
    '***********************************************************************************

    '***********************************************************************************
    '* davide 5/11/2010: proceso de lectura c�digo de barras de los tubos origen       *
    '*---------------------------------------------------------------------------------*
    '* genera un fichero bcdata.dat con los c�digos de barras de los tubos muestras    *
    '***********************************************************************************
	'START CODING HERE -----------------------------------------------------------------
	'Variable declaration
	'SUBPROCESS_ACTIVE = True
	'XToolApp.CallSub("C:\Lirix\data\process\Read_Tubes.bas","main")
    '***********************************************************************************
    '* davide 5/11/2010: proceso de lectura c�digo de barras de los tubos origen       *
    '*---------------------------------------------------------------------------------*
    '***********************************************************************************

    n = 1
	DestRackName(0) = "MP_001"
	DestRackName(1) = "MP_002"

	If Exists_File(fchMuestras) Then
		vNumMuestras = Lee_Numero_de_Muestras(fchMuestras)
	Else
		Err.Raise(-1571, "y-Plat Aliquot", "No se encuentra el fichero " & fchMuestras & " u el fichero es demasiado antiguo.")
	End If

	If Exists_File(fchplateBC) Then 'Hay que comprobar que el fichero no sea antiguo
		aNumPlacas(0) = Lee_Codigos_de_Placas(fchplateBC, DestRackName(0))
		aNumPlacas(1) = Lee_Codigos_de_Placas(fchplateBC, DestRackName(1))
	Else
		Err.Raise(-1571, "y-Plat Aliquot", "No se encuentra el fichero " & fchplateBC & " u el fichero es demasiado antiguo.")
	End If

	If Exists_File(fchbcdata) Then 'Hay que comprobar que el fichero no sea antiguo
		If vNumMuestras <> Lee_Numero_de_Muestras_enbcdata(fchbcdata) Then
			Err.Raise(-1572, "y-Plat Aliquot", "El n�mero de c�digos de barra en " & fchbcdata & " no coincide con el n�mero en " & fchMuestras)
		End If
	Else
		Err.Raise(-1571, "y-Plat Aliquot", "No se encuentra el fichero " & fchbcdata & " u el fichero es demasiado antiguo.")
	End If

	'DestRackName(2) = "MP_003"
	'DestRackName(3) = "MP_004"
	MAXDispensedTarget = 0
	MAXDispensedRack = 0
	OverAllDispensedQuotas = 0
	strFilePath = AppPath & "\log\INyDIA_Distribute_Log.MDB"

    '***********************************************************************************
    '* davide 5/11/2010:                                                               *
    '*---------------------------------------------------------------------------------*
    '* este debe llamarse como el rack de origen que se lee en el bcdata.dat           *
    '***********************************************************************************
	'CurrentTubeSet.SourceRackName = "SampleRack_001"
    '***********************************************************************************
    '* davide 5/11/2010:                                                               *
    '*---------------------------------------------------------------------------------*
    '***********************************************************************************

    'Set xvol = Robot.CreateVolArr(1)
	Call FillWellMatrix96pos(WellMatrix())
	Call Check_MDB_File(strFilePath)
	vExecution_ID = Insert_Execution(strFilePath)

	Set db = Dbs.OpenDatabase(strFilePath, False, False)

	For i = 1 To vNumMuestras
		BCLine = aa.INIGetValue (AppPath & "\data\bcdata.dat", "Barcode", CStr(i))
		BC = Split(BCLine, ",")
		BCNumber = BC(0)
		BCRack = BC(2)
		BCPos = CInt(BC(3))
		xs.WriteLog(LOG_INFO, BCLine)

		CurrentPos = MAXDispensedTarget + 1
		If CurrentPos > MAX_RACK_POS Then
			CurrentRack = MAXDispensedRack + Int(CurrentPos / MAX_RACK_POS)
			CurrentPos = CurrentPos Mod MAX_RACK_POS
		Else
			CurrentRack = MAXDispensedRack + 0
		End If

		Call InsertRecord(vExecution_ID, BCNumber, BCRack, BCPos, DestRackName(CurrentRack), CurrentPos, WellMatrix(CurrentPos), CurrentQuota, db)

		OverAllDispensedQuotas = OverAllDispensedQuotas + 1
		MAXDispensedTarget = OverAllDispensedQuotas Mod MAX_RACK_POS
		MAXDispensedRack = Int(OverAllDispensedQuotas / MAX_RACK_POS)
	Next i
	db.Close
	Set db = Nothing
	Set Dbs = Nothing

	'xs.DropTips(AllChannels, 0,False, YES,0, NO,YES,NO,YES,YES)

'    For i = 0 To ProcessSourceTubes -1 Step NumberOfTips
'		Set db = Dbs.OpenDatabase(strFilePath, False, False)
'
'        Call Initialize_TubeSet(CurrentTubeSet)
'		xs.GetTips(AllChannels, GT_DisplayError, Rainin_1000ul,"",0,1)
'		xs.DetectLiquid(CurrentTubeSet.SourceRackName, AllChannels, i+1, 1,, xvol,0.00)
'		xs.WriteLog(LOG_DEBUG, "L�quidos detectados: " & CStr(xvol.Vol(0)) & " - " & CStr(xvol.Vol(1)) &  " - " & CStr(xvol.Vol(2)) &  " - " & CStr(xvol.Vol(3)))
'
'		For a = 0 To 7
'			With CurrentTubeSet.Tubes(a)
'				.DetectedVolume = xvol.Vol(a)
'				.ReqQuotaVolume = VolOfQuota
'				.UseIncompleteQuota = IncompleteQuota
'				.MinLastQuota = MinLastQuota
'			End With
'		Next a
'
'		Call Calculate_Previous_Tube_Quota(CurrentTubeSet)
'		Call Calculate_TubeSetTotQuotas(CurrentTubeSet)
'		Call Log_TubeSet_Info(CurrentTubeSet)
'		On Error Resume Next
'			Kill PosInfoFile
'		On Error GoTo 0
'
'		ff.WriteINI(PosInfoFile, "Worklist", "Num", CStr(CurrentTubeSet.TotQuotas))
'		LineCounter = 1
'
'		Call Check_Empty_TubeSet(CurrentTubeSet)
'		Call Log_TubeSet_IsEmpty(CurrentTubeSet)
'
'		Do Until CurrentTubeSet.IsEmpty
'
'			For a = 0 To 7
'				With CurrentTubeSet.Tubes(a)
'					If Not .IsTubeEmpty And .RemainingVolume > MinLastQuota Then
'
'						CurrentPos = MAXDispensedTarget + .PreviousTubesQuotas + .NoOfDispensedQuotas + 1
'						If CurrentPos > MAX_RACK_POS Then
'							CurrentRack = MAXDispensedRack + Int(CurrentPos / MAX_RACK_POS)
'							CurrentPos = CurrentPos Mod MAX_RACK_POS
'						Else
'							CurrentRack = MAXDispensedRack + 0
'						End If
'
'						On Error GoTo ErrorInPosinfo
'						CurrentQuota = IIf(.RemainingVolume < .ReqQuotaVolume, .RemainingVolume -1, .ReqQuotaVolume)
'							PosInfoLine = ",," & _
'								CurrentTubeSet.SourceRackName & "," & _
'								i + .TubeOrderNo & ",," & _
'								DestRackName(CurrentRack) & "," & _
'								CurrentPos & "," & _
'								CurrentQuota & ","
'						On Error GoTo 0
'
'						Log_PosInfoLine(CStr(LineCounter) & "=" & PosInfoLine)
'						ff.WriteINI(PosInfoFile, "Worklist", CStr(LineCounter), PosInfoLine)
'						BCLine = aa.INIGetValue (AppPath & "\data\bcdata.dat", "Barcode", CStr(.TubeOrderNo))
'						BC = Split(BCLine, ",")
'						BCNumber = BC(0)
'						BCRack = BC(2)
'						BCPos = CInt(BC(3))
'						Debug.Print BCNumber
'
'						Call InsertRecord(vExecution_ID, BCNumber, CurrentTubeSet.SourceRackName, i + .TubeOrderNo, DestRackName(CurrentRack), CurrentPos, WellMatrix(CurrentPos), CurrentQuota, db)
'						LineCounter = LineCounter + 1
'						.NoOfDispensedQuotas = .NoOfDispensedQuotas + 1
'						OverAllDispensedQuotas = OverAllDispensedQuotas + 1
'					ElseIf .TubeOrderNo <= NumberOfTips Then
'						PosInfoLine = ",,,,,,,,"
'						Log_PosInfoLine(CStr(LineCounter) & "=" & PosInfoLine)
'						ff.WriteINI(PosInfoFile, "Worklist", CStr(LineCounter), PosInfoLine)
'						LineCounter = LineCounter + 1
'					End If
'				End With
'			Next a
'
'			Call Check_Empty_TubeSet(CurrentTubeSet)
'		Loop
'		ff.WriteINI(PosInfoFile, "Worklist", "Num", CStr(LineCounter-1))
'
'
'OutOfBlue:
'	MAXDispensedTarget = OverAllDispensedQuotas Mod MAX_RACK_POS
'	MAXDispensedRack = Int(OverAllDispensedQuotas / MAX_RACK_POS)
'
'	Call Check_Empty_TubeSet(CurrentTubeSet)
'	Call Log_TubeSet_IsEmpty(CurrentTubeSet)
'	Call Empty_Tubeset(CurrentTubeSet)
'	db.Close
'

    '***********************************************************************************
    '* davide 5/11/2010: proceso de lectura distribuci�n del liquido                   *
    '*---------------------------------------------------------------------------------*
    '* aqu� se lanza la macro para distribuir las muestras 1 a 1                       *
    '*---------------------------------------------------------------------------------*
    '***********************************************************************************
    '* �MUY IMPORTANTE! SUBPROCESS_ACTIVE = True �MUY IMPORTANTE!                      *
    '***********************************************************************************
	'START CODING HERE -----------------------------------------------------------------
	'Dim Par1 As Variant, Par2 As Variant, Par3 As Variant
	'Dim FunctName$
	'SUBPROCESS_ACTIVE = True
	'FunctName = ""
	'If FunctName="" Then
	'	FunctName="main"
	'End If

	'If "" <> "" Then
		'Par1=CVar("")
		'If "" <> "" Then
			'Par2=CVar("")
			'If "" <> "" Then
				'Par3=CVar("")
				'XToolApp.CallSub("C:\Lirix\data\process\TEST_002.bas",FunctName,Par1,Par2,Par3)
			'Else
				'XToolApp.CallSub("C:\Lirix\data\process\TEST_002.bas",FunctName,Par1,Par2)
			'End If
		'Else
			'XToolApp.CallSub("C:\Lirix\data\process\TEST_002.bas",FunctName,Par1)
		'End If
	'Else
	'	XToolApp.CallSub("C:\Lirix\data\process\TEST_002.bas","main")
	'End If
    '***********************************************************************************
    '* davide 5/11/2010: proceso de lectura distribuci�n del liquido                   *
    '***********************************************************************************
    '***********************************************************************************
    '* �MUY IMPORTANTE! SUBPROCESS_ACTIVE = True �MUY IMPORTANTE!                      *
    '***********************************************************************************

'	n= n + NumberOfTips
'	xs.DropTips(AllChannels, 0, NO, NO,0, NO,YES,NO,YES,YES)

'Next i

StrXMLFileName = Create_XML_File(strFilePath, "Muestras del "& Format(Now,"dd mmmm yyyy hh_nn_ss") & ".XML", vExecution_ID)

Call Execute_HTML_File(StrXMLFileName)

xs.MoveAbsPos (HomePosition)


        'END CODING  ---------------------------------------------------------------------------------
    End If

    Call Cleanup
Exit Sub

ErrorInPosinfo:
Select Case Err.Number
	Case 10023
		MsgBox ("There's no space left on destination racks, reduce number of sources")
		Err.Clear
		On Error GoTo 0
	Case Else
		MsgBox("Oooops, there's something wrong here")
		Err.Raise Err.Number
End Select
        xs.MoveAbsPos (HomePosition)
    Call Cleanup

End Sub


Sub InsertRecord(vExecution_ID As Long, vBarCode As String, vSourceRack As String, vSourceTube As Integer, vTargetRack As String, vWellNumber As Integer, vPosition As String, vCurrentQuota As Long, db As DAO.Database)
Dim rs As DAO.Recordset

	Set rs = db.OpenRecordset("Distribution_Lines")

	'Add a new record
	rs.AddNew
	rs!Execution_ID = vExecution_ID
	If vBarCode = "" Then
		rs!BarCode = 0
	Else
		rs!BarCode = vBarCode
	End If
	rs!SourceRack = vSourceRack
	rs!SourceTube = vSourceTube
	rs!TargetRack = vTargetRack
	rs!WellNumber = vWellNumber
	rs!Position = vPosition
	rs!QuotaVolume = vCurrentQuota

	rs.Update

	Set rs = Nothing

End Sub
Sub FillWellMatrix96pos(WellMatrix() As String)

      WellMatrix(0)="!!" 'No existe en la placa la posici�n 0

      WellMatrix(1)="A1"
      WellMatrix(2)="B1"
      WellMatrix(3)="C1"
      WellMatrix(4)="D1"
      WellMatrix(5)="E1"
      WellMatrix(6)="F1"
      WellMatrix(7)="G1"
      WellMatrix(8)="H1"
      WellMatrix(9)="A2"
      WellMatrix(10)="B2"
      WellMatrix(11)="C2"
      WellMatrix(12)="D2"
      WellMatrix(13)="E2"
      WellMatrix(14)="F2"
      WellMatrix(15)="G2"
      WellMatrix(16)="H2"
      WellMatrix(17)="A3"
      WellMatrix(18)="B3"
      WellMatrix(19)="C3"
      WellMatrix(20)="D3"
      WellMatrix(21)="E3"
      WellMatrix(22)="F3"
      WellMatrix(23)="G3"
      WellMatrix(24)="H3"
      WellMatrix(25)="A4"
      WellMatrix(26)="B4"
      WellMatrix(27)="C4"
      WellMatrix(28)="D4"
      WellMatrix(29)="E4"
      WellMatrix(30)="F4"
      WellMatrix(31)="G4"
      WellMatrix(32)="H4"
      WellMatrix(33)="A5"
      WellMatrix(34)="B5"
      WellMatrix(35)="C5"
      WellMatrix(36)="D5"
      WellMatrix(37)="E5"
      WellMatrix(38)="F5"
      WellMatrix(39)="G5"
      WellMatrix(40)="H5"
      WellMatrix(41)="A6"
      WellMatrix(42)="B6"
      WellMatrix(43)="C6"
      WellMatrix(44)="D6"
      WellMatrix(45)="E6"
      WellMatrix(46)="F6"
      WellMatrix(47)="G6"
      WellMatrix(48)="H6"
      WellMatrix(49)="A7"
      WellMatrix(50)="B7"
      WellMatrix(51)="C7"
      WellMatrix(52)="D7"
      WellMatrix(53)="E7"
      WellMatrix(54)="F7"
      WellMatrix(55)="G7"
      WellMatrix(56)="H7"
      WellMatrix(57)="A8"
      WellMatrix(58)="B8"
      WellMatrix(59)="C8"
      WellMatrix(60)="D8"
      WellMatrix(61)="E8"
      WellMatrix(62)="F8"
      WellMatrix(63)="G8"
      WellMatrix(64)="H8"
      WellMatrix(65)="A9"
      WellMatrix(66)="B9"
      WellMatrix(67)="C9"
      WellMatrix(68)="D9"
      WellMatrix(69)="E9"
      WellMatrix(70)="F9"
      WellMatrix(71)="G9"
      WellMatrix(72)="H9"
      WellMatrix(73)="A10"
      WellMatrix(74)="B10"
      WellMatrix(75)="C10"
      WellMatrix(76)="D10"
      WellMatrix(77)="E10"
      WellMatrix(78)="F10"
      WellMatrix(79)="G10"
      WellMatrix(80)="H10"
      WellMatrix(81)="A11"
      WellMatrix(82)="B11"
      WellMatrix(83)="C11"
      WellMatrix(84)="D11"
      WellMatrix(85)="E11"
      WellMatrix(86)="F11"
      WellMatrix(87)="G11"
      WellMatrix(88)="H11"
      WellMatrix(89)="A12"
      WellMatrix(90)="B12"
      WellMatrix(91)="C12"
      WellMatrix(92)="D12"
      WellMatrix(93)="E12"
      WellMatrix(94)="F12"
      WellMatrix(95)="G12"
      WellMatrix(96)="H12"

End Sub
Sub Create_MDB_File(strFilePath As String)
	Const dbInteger As Integer = 3
	Const dbText As Integer = 10
	Const dbLong As Integer = 4
	Const dbAutoIncrField As Integer = 16
	Const dbDate As Integer = 8
	Const dbMemo As Integer = 12
	Const dbRelationUpdateCascade As Integer = 256
	Const dbRelationDeleteCascade As Integer = 4096
Dim Dbs As DAO.DBEngine
Dim db As DAO.Database
Dim Tbl As DAO.TableDef
Dim Fld As DAO.Field
Dim ind As DAO.Index
Const dbLangSpanish = ";LANGID=0x0409;CP=1252;COUNTRY=0" 'DAO.LanguageConstants

	Set Dbs = CreateObject("DAO.DBEngine.36")
	Set db = Dbs.CreateDatabase(strFilePath, dbLangSpanish)

	'Create a new table definition for a table called Executions
	Set Tbl = db.CreateTableDef("Executions")
	Set Fld = Tbl.CreateField("Execution_ID", dbLong)
   	Fld.attributes = dbAutoIncrField
	Tbl.Fields.Append Fld

	Set Fld = Tbl.CreateField("Execution_Date", dbDate)
	Tbl.Fields.Append Fld
	Set Fld = Tbl.CreateField("Execution_Notes", dbMemo)
	Tbl.Fields.Append Fld

	db.TableDefs.Append Tbl

    Set ind = Tbl.CreateIndex("PK_Executions")
    With ind
        .Fields.Append .CreateField("Execution_ID")
        .Unique = True
        .Primary = True
    End With
    Tbl.Indexes.Append ind
    'Refresh the display of this collection.
    Tbl.Indexes.Refresh

    'Clean up
    Set ind = Nothing

	'Create a new table definition for a table called Distribution_Lines
	Set Tbl = db.CreateTableDef("Distribution_Lines")

	'Create a new field in NewTable and call it NewField
	Set Fld = Tbl.CreateField("Line_ID", dbLong)
   	Fld.attributes = dbAutoIncrField
	Tbl.Fields.Append Fld
	Set Fld = Tbl.CreateField("Execution_ID", dbLong)
	Fld.Required = True
	Tbl.Fields.Append Fld
	Set Fld = Tbl.CreateField("BarCode", dbText)
	Tbl.Fields.Append Fld
	Set Fld = Tbl.CreateField("SourceRack", dbText)
	Tbl.Fields.Append Fld
	Set Fld = Tbl.CreateField("SourceTube", dbInteger)
	Tbl.Fields.Append Fld
	Set Fld = Tbl.CreateField("TargetRack", dbText)
	Tbl.Fields.Append Fld
	Set Fld = Tbl.CreateField("Position", dbText)
	Tbl.Fields.Append Fld
	Set Fld = Tbl.CreateField("WellNumber", dbInteger)
	Tbl.Fields.Append Fld
	Set Fld = Tbl.CreateField("QuotaVolume", dbLong)
	Tbl.Fields.Append Fld

	db.TableDefs.Append Tbl
    Set ind = Tbl.CreateIndex("PK_Distribution_Lines")
    With ind
        .Fields.Append .CreateField("Line_ID")
        .Unique = True
        .Primary = True
    End With
    Tbl.Indexes.Append ind
    Set ind = Tbl.CreateIndex("FK_Distribution_Lines_Executions")
    With ind
        .Fields.Append .CreateField("Execution_ID")
        .Unique = False
        .Primary = False
    End With
    Tbl.Indexes.Append ind
    Tbl.Indexes.Refresh

    'Clean up
    Set ind = Nothing

 	Dim rel As DAO.Relation
    Set rel = db.CreateRelation("Executions_Distribution_Lines")
    With rel
        'Specify the primary table.
        .Table = "Executions"
        'Specify the related table.
        .ForeignTable = "Distribution_Lines"
        'Specify attributes for cascading updates and deletes.
        .attributes = dbRelationUpdateCascade + dbRelationDeleteCascade

        'Add the fields to the relation.
        'Field name in primary table.
        Set Fld = .CreateField("Execution_ID")
        'Field name in related table.
        Fld.ForeignName = "Execution_ID"
        'Append the field.
        .Fields.Append Fld

        'Repeat for other fields if a multi-field relation.
    End With
    db.Relations.Append rel

    Dim qdf As DAO.QueryDef

    'Set db = CurrentDb()

    'The next line creates and automatically appends the QueryDef.
    Set qdf = db.CreateQueryDef("qry_ExportHTML")

    'Set the SQL property to a string representing a SQL statement.
    qdf.SQL = "SELECT [Execution_ID], [BarCode], [SourceRack], [SourceTube], [TargetRack], [Position], [WellNumber], [QuotaVolume] FROM Distribution_Lines  ORDER BY [TargetRack], [WellNumber];"

    'Do not append: QueryDef is automatically appended!

    Set qdf = Nothing

	db.Close

	Set rel = Nothing

	Set db = Nothing
	Set Dbs = Nothing

End Sub
Function Create_XML_File(strFilePath As String, StrXMLFileName As String, vExecution_ID) As String
Dim varPathCurrent As String
Dim filesys As Object
Dim xmlDoc As New DOMDocument
Dim conn As Object, rts As Object

	Set filesys = CreateObject("Scripting.FileSystemObject")
	varPathCurrent = filesys.GetParentFolderName(strFilePath)
	Set filesys = Nothing

	Call Delete_File(varPathCurrent & "\" & StrXMLFileName)
	Set conn = CreateObject("ADODB.Connection")
	conn.Provider = "Microsoft.Jet.OLEDB.4.0"
	Call conn.open(strFilePath)

	Set rts = CreateObject("ADODB.recordset")
	'Call rts.open("SELECT [SourceRack] AS Origen, [SourceTube], [TargetRack] As destino, [Position], [WellNumber], [QuotaVolume] FROM qry_ExportHTML WHERE [Execution_ID]= " & vExecution_ID, conn)
	Call rts.open("SELECT [TargetRack] AS [Destination Plate ID], [BarCode] AS [Source Sample ID], [SourceTube] AS [Posici�n origen], [Position] AS [Posici�n destino] FROM qry_ExportHTML WHERE [Execution_ID]= " & vExecution_ID, conn)
	'Save the Recordset into a DOM tree
	Call rts.Save(xmlDoc, 1)
	Call xmlDoc.insertBefore(xmlDoc.createProcessingInstruction("xml-stylesheet", "type=""text/xsl"" href=""INyDIA_Distribute_Log.XSL"""), xmlDoc.documentElement)

	'Writes the datetime of the creation
	Dim xmlFechaNode As IXMLDOMNode
	Set xmlFechaNode = xmlDoc.documentElement.appendChild(xmlDoc.createNode(NODE_ELEMENT, "fecha_hora", ""))
	xmlFechaNode.text = Format(Now, "dddd dd mmmm - hh:nn")

	Call xmlDoc.Save(varPathCurrent & "\" & StrXMLFileName)
	Set rts = Nothing
	Set conn = Nothing
	Set xmlDoc = Nothing
	Create_XML_File = varPathCurrent & "\" & StrXMLFileName

End Function
Sub Clean_HTML_File(strHTMFileName As String)

	Call Replace_String(" DIR=LTR", "", strHTMFileName)
	Call Replace_String("charset=Windows-1252"">", "charset=Windows-1252""/>", strHTMFileName)
	Call Replace_String("ALIGN=LEFT", "ALIGN=""LEFT""", strHTMFileName)
	Call Replace_String("ALIGN=RIGHT", "ALIGN=""RIGHT""", strHTMFileName)
	Call Replace_String(" BORDER", " BORDER=""1""", strHTMFileName)
	Call Replace_String("</TR>" & vbCrLf & "<TD", "</TR>" & vbCrLf & "<TR>" & vbCrLf & "<TD", strHTMFileName)

End Sub

Sub Delete_File(strFilePath As String)
Dim objFSO As Object

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If (objFSO.FileExists(strFilePath)) Then
    	objFSO.DeleteFile(strFilePath)
	End If
	Set objFSO = Nothing

End Sub
Sub Execute_HTML_File(strHTMFileName As String)
Dim Shl

	Set Shl = CreateObject("WScript.Shell")
    Shl.Run Chr(34) & strHTMFileName & Chr(34), 1, False
    Set Shl = Nothing
    'X = Shell(strHTMFileName)

End Sub
Sub Replace_String(Find As String, ReplaceWith As String, FileName As String)
Dim dFileContents As String, fileContents As String

	fileContents = GetFile(FileName)
	dFileContents = Replace(fileContents, Find, ReplaceWith, 1, -1)

	If dFileContents <> fileContents Then
	 'write result If different
	 WriteFile FileName, dFileContents

	 'Wscript.Echo "Replace done."
	 If Len(ReplaceWith) <> Len(Find) Then 'Can we count n of replacements?
	   'Wscript.Echo  ( (Len(dFileContents) - Len(FileContents)) / (Len(ReplaceWith)-Len(Find)) ) & " replacements."
	 End If
	Else
	 Debug.Print  "Searched string Not In the source file"
	End If

End Sub

Function GetFile(FileName As String)
Dim FS As Object
	 If FileName<>"" Then
	   Dim FileStream As Object
	   Set FS = CreateObject("Scripting.FileSystemObject")
	     On Error Resume Next
	     Set FileStream = FS.OpenTextFile(FileName)
	     GetFile = FileStream.ReadAll
	 End If
	 Set FileStream = Nothing
	 Set FS = Nothing
End Function

'Write string As a text file.
Function WriteFile(FileName As String, Contents As String)
Dim OutStream, FS As Object

	 On Error Resume Next
	 Set FS = CreateObject("Scripting.FileSystemObject")
	   Set OutStream = FS.OpenTextFile(FileName, 2, True)
	   OutStream.Write Contents
     Set OutStream =Nothing
	 Set FS = Nothing

End Function
Sub Check_MDB_File(strFilePath As String)
	Dim objFSO As Object
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not (objFSO.FileExists(strFilePath)) Then
    	Create_MDB_File(strFilePath)
	End If
	Set objFSO = Nothing

End Sub
Function Insert_Execution(strFilePath As String) As Long
	Dim Dbs As DAO.DBEngine
	Dim db As DAO.Database
	Dim rs As DAO.Recordset
	Dim vExecution_ID As Long

	Set Dbs = CreateObject("DAO.DBEngine.36")
	Set db = Dbs.OpenDatabase(strFilePath, False, False)

	Set rs = db.OpenRecordset("Executions")
	rs.AddNew
	rs!Execution_Date = Date
	rs!Execution_Notes = "y-Plat Aliq: distribuci�n 1 a 1"
	vExecution_ID = rs!Execution_ID
	rs.Update
	Set rs = Nothing
	db.Close
	Set db = Nothing
	Set Dbs = Nothing
	Insert_Execution = vExecution_ID

End Function
Sub Beautify_XML_File(strHTMFileName As String)
Dim xDoc As New MSXML.DOMDocument

      Set xDoc = CreateObject("Microsoft.XMLDOM")
      If xDoc.Load(strHTMFileName) Then
              'Debug.Print "OK"
              'Dim xmlProcessInstruction As IXMLDOMProcessingInstruction

              'Set xmlProcessInstruction = xDoc.createProcessingInstruction("xml-stylesheet", "type=""text/xsl"" href=""INyDIA_Distribute_Log.XSL""")
              'Set HEADNode = xDoc.selectSingleNode("/HTML/HEAD")
              'Set element = xDoc.createElement("LINK") '<LINK href="style.css" rel="stylesheet" Type="text/css">
              'Set objhrefAttr = xDoc.createAttribute("href")
              'Set objrelAttr = xDoc.createAttribute("rel")
              'Set objtypeAttr = xDoc.createAttribute("type")
              'element.setAttribute "href", "style.css"
              'element.setAttribute "rel", "stylesheet"
              'element.setAttribute "type", "text/css"

              'HEADNode.appendChild (element)
              'Set objhrefAttr = Nothing
              'Set objrelAttr = Nothing
              'Set objtypeAttr = Nothing
              'Set element = Nothing
              'Set HEADNode = Nothing

              'Set TABLENode = xDoc.selectSingleNode("/HTML/BODY/TABLE")
              'Set objidAttr = xDoc.createAttribute("id")
              'TABLENode.setAttribute "id", "gradient-style"
              'Set objidAttr = Nothing
              'Set TABLENode = Nothing

              'xDoc.insertBefore (xmlProcessInstruction)
              xDoc.insertBefore xDoc.createProcessingInstruction("xml-stylesheet", "type=""text/xsl"" href=""INyDIA_Distribute_Log.XSL"""), xDoc.documentElement
              'xDoc.appendChild

              xDoc.Save strHTMFileName
      Else
         ' The document failed to load.
         Dim strErrText 'As String
         Dim xPE 'As MSXML.IXMLDOMParseError
         ' Obtain the ParseError object
         Set xPE = xDoc.parseError
         With xPE
            strErrText = "Your XML Document failed to load" & _
              "due the following error." & vbCrLf & _
              "Error #: " & .errorCode & ": " & xPE.reason & _
              "Line #: " & .line & vbCrLf & _
              "Line Position: " & .linepos & vbCrLf & _
              "Position In File: " & .filepos & vbCrLf & _
              "Source Text: " & .srcText & vbCrLf & _
              "Document URL: " & .url
          End With
         Set xPE = Nothing

          Debug.Print strErrText
      End If

      Set xDoc = Nothing

End Sub

Function Exists_File(strFilePath As String) As Boolean
Dim objFSO As Object, LastModifiedDate As Date
'Hay que comprobar que el fichero no sea antiguo

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If (objFSO.FileExists(strFilePath)) Then
		LastModifiedDate = objFSO.GetFile(strFilePath).DateLastModified

		If DateDiff("h", LastModifiedDate, Now) > 1 Then
			Dim MsgRes As VbMsgBoxResult
			MsgRes = MsgBox(strFilePath & " fue modificado hace " & DateDiff("h", LastModifiedDate, Now) & " horas. �Continuar?" ,vbOkCancel,"La fecha/hora del fichero es demasiado antigua")
			If MsgRes=vbCancel Then
				xs.AbortRunDetected = True
				Exists_File = False
			Else
				Exists_File = True
			End If
		Else
			Exists_File = True
		End If
	Else
		Exists_File = False
	End If
	Set objFSO = Nothing

End Function

Function Lee_Numero_de_Muestras(fchMuestras As String) As Integer
'Dim ff

  	'Set ff = CreateObject("cuf.FileFunctions")

	Lee_Numero_de_Muestras = CInt(ff.GetINIString(fchMuestras, "MUESTRAS", "NoMuestras"))
	'Set ff = Nothing
End Function

Function Lee_Codigos_de_Placas(fchplateBC As String, MP As String) As Long
'Dim ff

  	'Set ff = CreateObject("cuf.FileFunctions")

	Lee_Codigos_de_Placas = CLng(ff.GetINIString(fchplateBC, "PLATES BC", MP))
	'Set ff = Nothing
End Function

Function Lee_Numero_de_Muestras_enbcdata(fchbcdata As String) As Integer
'Dim ff

  	'Set ff = CreateObject("cuf.FileFunctions")

	Lee_Numero_de_Muestras_enbcdata = CInt(ff.GetINIString(fchbcdata, "Barcode", "Num"))
	'Set ff = Nothing
End Function
