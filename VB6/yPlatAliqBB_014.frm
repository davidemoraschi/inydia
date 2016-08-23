VERSION 5.00
Begin VB.Form yPlatAliqBB_014 
   Caption         =   "yPlatAliqBB_014"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Width           =   1215
   End
End
Attribute VB_Name = "yPlatAliqBB_014"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const fldLirix = "C:\Lirix"                             'Carpeta raiz de Lirix
Const fchMuestras = fldLirix & "\data\muestras.dat"     'Fichero con el número de muestras
Const fchplateBC = fldLirix & "\data\plateBC.dat"       'Fichero con códigos de barras de las placas destino
Const fchbcdata = fldLirix & "\data\bcdata.dat"         'Fichero con códigos de barras de las muestras
Const strFilePath = fldLirix & "\data\report\INyDIA_Distribute_Log.MDB"
Const INI_FILE_NAME = fldLirix & "\data\macros\yPlatAliqBB_014.ini"    'Fichero de configuración
Const PosInfoFile = fldLirix & "\data\piplist\posinfo.dat" '+ Format(Now,"yyyymmdd-hhnnss")+".dat"
Const MAX_RACK_POS = 96

Public ProcessLayout$
Public NumberOfTips As Integer '= 8
Public ProcessSourceTubes As Integer  '= 24
Public VolOfQuota As Long  '= 500
Public IncompleteQuota As Boolean  '= True  ' ***** Propiedad de la clase INyDIA_SourceTube
Public MinLastQuota As Long '= 100          ' ***** Propiedad de la clase INyDIA_SourceTube
Public SourceRackPref As String
Public DestRackPref As String
Public DestPlateNum As Long
Public DestReptFold As String
Public DestResFold As String
Function InitGlobalVariables() As Boolean

'Set default values
    'ProcessSourceTubes  = 192
    'VolOfQuota = 500
   ' ProcessSourceTubes = Lee_Numero_de_Muestras(fchMuestras)

'INPUTERROR:
'    Begin Dialog UserDialog 430,154,"Insert Process Varibales" ' %GRID:10,7,1,1
'        OKButton 310, 119, 100, 21
'        CancelButton 190, 119, 100, 21
'        Text 10, 35, 140, 21, "No of Source tubes"
'        Text 10, 63, 100, 28, "Volumen Alíquota"
'        TextBox 180, 35, 60, 21, .ProcessSourceTubes
'        TextBox 110, 70, 60, 21, .VolOfQuota
'        Text 60, 7, 300, 21, "Inicio", .header, 2
'        Text 10, 98, 60, 21, "NOTA:", .Text1
'        Text 80, 98, 350, 14, "Volumen en microlitros", .Text2
'    End Dialog

'    Dim Dlg As UserDialog

    On Error GoTo CANCELPRESSED

'   Dlg.ProcessSourceTubes$ = Str$(ProcessSourceTubes)
'    Dlg.VolOfQuota$ = Str$(VolOfQuota)

    '**************************************************************************
    '* davide 29/11/2010:deshabilita la dialog box de parámetros al principio *
    '*------------------------------------------------------------------------*
    '**************************************************************************
    'Dialog Dlg
    '**************************************************************************
    '* davide 29/11/2010:deshabilita la dialog box de parámetros al principio *
    '*------------------------------------------------------------------------*
    '**************************************************************************

    'ProcessSourceTubes  = CLng(Dlg.ProcessSourceTubes$)
    'If ProcessSourceTubes  < 0 Or NoOfCycles > 10 Then
     '   MsgBox "Invalid Value for 'NoOfCycles', Allowed Range: [0,10]"
      '  GoTo INPUTERROR
    'End If
    'VolOfQuota = CLng(Dlg.VolOfQuota$)

    InitGlobalVariables = True
    Exit Function

CANCELPRESSED:
    InitGlobalVariables = False

End Function
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
    Call xs.SetOutputSwitch(ResSwitch1, , , S_ON, "Switch blue diodes On")
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

    Call xs.SetOutputSwitch(ResSwitch1, , , S_OFF, "Switch blue diodes Off")

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

Private Sub Command1_Click()
    Static MAXDispensedTarget As Long
    Static MAXDispensedRack As Integer
    Static OverAllDispensedQuotas As Long

    Dim Dbs As New DAO.DBEngine, db As DAO.Database
    Dim InitVar As Boolean
    Dim CurrentTubeSet As SourceTubeSet, xvol As SamVolArr
    Dim DestRackName() As String, WellMatrix(96) As String, BC() As String
    Dim a As Integer, CurrentPos As Integer, CurrentRack As Integer
    Dim i As Long, LineCounter As Long, n As Long, vExecution_ID As Long, CurrentQuota As Long
    Dim PosInfoLine As String, StrXMLFileName As String, BCLine As String, BCNumber As String, BCRack As String, BCPos As Integer

    Dim SrcAssName As Long  'sirve?
    'Dim vNumMuestras As Integer' vExecution_ID As Long, ,
    Dim aNumPlacas() As String

    If Exists_FileINI(INI_FILE_NAME) Then
        ProcessLayout = Lee_Fichero_INI_str(INI_FILE_NAME, "ProcessLayout", "Layour de Proceso")
    Else
        Call Err.Raise(-1571, "y-Plat Aliquot", "No se encuentra el fichero " & INI_FILE_NAME)
    End If

    Call CreateXToolObject(ProcessLayout)

    InitVar = InitGlobalVariables
    If InitVar = True Then
        Call Init
        Call Start(ProcessLayout)


        'START CODING HERE      --------------------------------------------------------------------------
        'Variable declaration
    SUBPROCESS_ACTIVE = True
    Call XToolApp.CallSub("C:\Lirix\data\process\Barcodes.bas", "main")
    SUBPROCESS_ACTIVE = False

    Debug.Print "Reading INI file" & String$(20 - Len("Reading INI file: "), "_") & ":" & INI_FILE_NAME

    If Exists_FileINI(INI_FILE_NAME) Then
        VolOfQuota = CLng(Lee_Fichero_INI_str(INI_FILE_NAME, "VolOfQuota", "Volumen de alícuota"))
        MinLastQuota = CLng(Lee_Fichero_INI_str(INI_FILE_NAME, "MinLastQuota", "Volumen mínimo"))
        NumberOfTips = CInt(Lee_Fichero_INI_str(INI_FILE_NAME, "NumberOfTips", "Número de puntas"))
        SourceRackPref = Lee_Fichero_INI_str(INI_FILE_NAME, "SourceRackPref", "Prefijo para rack origen")
        DestRackPref = Lee_Fichero_INI_str(INI_FILE_NAME, "DestRackPref", "Prefijo para rack destino")
        DestPlateNum = CLng(Lee_Fichero_INI_str(INI_FILE_NAME, "DestPlateNum", "Número de placas destino"))
        DestReptFold = Lee_Fichero_INI_str(INI_FILE_NAME, "DestReptFold", "Carpeta para Reporte")
        DestResFold = Lee_Fichero_INI_str(INI_FILE_NAME, "DestResFold", "Carpeta para Resultados")

    If Exists_File(fchMuestras) Then
        ProcessSourceTubes = Lee_Numero_de_Muestras(fchMuestras)
        Debug.Print "ProcessSourceTubes" & String$(20 - Len("ProcessSourceTubes"), "_") & ":" & vbTab & ProcessSourceTubes
    Else
        ProcessSourceTubes = CInt(Lee_Fichero_INI_str(INI_FILE_NAME, "ProcessSourceTubes", "Número de tubos de origen"))
    End If

        IncompleteQuota = CBool(Lee_Fichero_INI_str(INI_FILE_NAME, "IncompleteQuota", "Utilizar última alícuota (true/false)?"))
    Else
        Call Err.Raise(-1571, "y-Plat Aliquot", "No se encuentra el fichero " & INI_FILE_NAME)
        '& " u el fichero es demasiado antiguo.")
    End If


    n = 1
'   DestRackName(0) = "MP_001"
'   DestRackName(1) = "MP_002"
'   DestRackName(2) = "MP_003"
'   DestRackName(3) = "MP_004"

    'Asigna un nombre a cada rack destino
    ReDim DestRackName(DestPlateNum - 1)
    ReDim aNumPlacas(DestPlateNum - 1)

    If Exists_File(fchplateBC) Then 'Hay que comprobar que el fichero no sea antiguo
        For i = 0 To DestPlateNum - 1
            DestRackName(i) = DestRackPref & Format(i + 1, "000")
            aNumPlacas(i) = Lee_Codigos_de_Placas(fchplateBC, DestRackName(i))
        Debug.Print "DestRackName" & String$(20 - Len("DestRackName"), "_") & ":" & vbTab & DestRackName(i) & "=" & aNumPlacas(i)
        Next
'       aNumPlacas(1) = Lee_Codigos_de_Placas(fchplateBC, DestRackName(1))
    Else
        Call Err.Raise(-1571, "y-Plat Aliquot", "No se encuentra el fichero " & fchplateBC & " u el fichero es demasiado antiguo.")
    End If
    If Exists_File(fchbcdata) Then 'Hay que comprobar que el fichero no sea antiguo
        If ProcessSourceTubes <> Lee_Numero_de_Muestras_enbcdata(fchbcdata) Then
            Call Err.Raise(-1572, "y-Plat Aliquot", "El número de códigos de barra en " & fchbcdata & " no coincide con el nímero de muestras.")
        End If
    Else
        Call Err.Raise(-1571, "y-Plat Aliquot", "No se encuentra el fichero " & fchbcdata & " u el fichero es demasiado antiguo.")
    End If

    MAXDispensedTarget = 0
    MAXDispensedRack = 0
    OverAllDispensedQuotas = 0
    'strFilePath = AppPath & "\log\INyDIA_Distribute_Log.MDB"
    CurrentTubeSet.SourceRackName = SourceRackPref & "S1"
    Set xvol = Robot.CreateVolArr(1)
    Call FillWellMatrix(WellMatrix())
    Call Check_File(strFilePath)
    vExecution_ID = Insert_Execution(strFilePath)

    'xs.DropTips(AllTips, 1,NO, NO,0,YES,YES,NO,YES,YES)

    For i = 0 To ProcessSourceTubes - 1 Step NumberOfTips
        Set db = Dbs.OpenDatabase(strFilePath, False, False)

        Call Initialize_TubeSet(CurrentTubeSet)
        Call xs.GetTips(AllTips, GT_DisplayError, Rainin_1000ul, "", 0, 1)
        Call xs.DetectLiquid(CurrentTubeSet.SourceRackName, 15, i + 1, 1, , xvol, 0#)
        Call xs.WriteLog(LOG_DEBUG, "Líquidos detectados: " & CStr(xvol.Vol(0)) & " - " & CStr(xvol.Vol(1)) & " - " & CStr(xvol.Vol(2)) & " - " & CStr(xvol.Vol(3)))

        'Call Generate_RNDTubeset (CurrentTubeSet, True, 600)

        For a = 0 To 7
            With CurrentTubeSet.Tubes(a)
                .DetectedVolume = xvol.Vol(a)
                .ReqQuotaVolume = VolOfQuota
                .UseIncompleteQuota = IncompleteQuota
                .MinLastQuota = MinLastQuota
            End With
        Next a


        Call Calculate_Previous_Tube_Quota(CurrentTubeSet)
        Call Calculate_TubeSetTotQuotas(CurrentTubeSet)
        'Call Log_TubeSet_Info(CurrentTubeSet)
        On Error Resume Next
            Kill PosInfoFile
        On Error GoTo 0

        Call ff.WriteINI(PosInfoFile, "Worklist", "Num", CStr(CurrentTubeSet.TotQuotas))
        LineCounter = 1

        Call Check_Empty_TubeSet(CurrentTubeSet)
        Call Log_TubeSet_IsEmpty(CurrentTubeSet)

        Do Until CurrentTubeSet.IsEmpty

            For a = 0 To 7
                With CurrentTubeSet.Tubes(a)
                    If Not .IsTubeEmpty And .RemainingVolume > MinLastQuota Then

                        CurrentPos = MAXDispensedTarget + .PreviousTubesQuotas + .NoOfDispensedQuotas + 1
                        If CurrentPos > MAX_RACK_POS Then
                            CurrentRack = MAXDispensedRack + Int(CurrentPos / MAX_RACK_POS)
                            CurrentPos = CurrentPos Mod MAX_RACK_POS
                        Else
                            CurrentRack = MAXDispensedRack + 0
                        End If

                        On Error GoTo ErrorInPosinfo
                        CurrentQuota = IIf(.RemainingVolume < .ReqQuotaVolume, .RemainingVolume - 1, .ReqQuotaVolume)
                            PosInfoLine = ",," & _
                                CurrentTubeSet.SourceRackName & "," & _
                                i + .TubeOrderNo & ",," & _
                                DestRackName(CurrentRack) & "," & _
                                CurrentPos & "," & _
                                CurrentQuota & ","
                        On Error GoTo 0

                        Log_PosInfoLine (CStr(LineCounter) & "=" & PosInfoLine)
                        Call ff.WriteINI(PosInfoFile, "Worklist", CStr(LineCounter), PosInfoLine)
                        BCLine = aa.INIGetValue(AppPath & "\data\bcdata.dat", "Barcode", CStr(i + .TubeOrderNo))
                        BC = Split(BCLine, ",")
                        BCNumber = BC(0)
                        BCRack = BC(2)
                        BCPos = CInt(BC(3))
                        Debug.Print BCNumber

                        Call Insert_Record(vExecution_ID, BCNumber, CurrentTubeSet.SourceRackName, i + .TubeOrderNo, DestRackName(CurrentRack), aNumPlacas(CurrentRack), CurrentPos, WellMatrix(CurrentPos), CurrentQuota, db)
                        'Call InsertRecord(vExecution_ID, BCNumber, BCRack, BCPos, DestRackName(CurrentRack), aNumPlacas(CurrentRack), CurrentPos, WellMatrix(CurrentPos), VolOfQuota, db)
                        LineCounter = LineCounter + 1
                        .NoOfDispensedQuotas = .NoOfDispensedQuotas + 1
                        OverAllDispensedQuotas = OverAllDispensedQuotas + 1
                    ElseIf .TubeOrderNo <= NumberOfTips Then
                        PosInfoLine = ",,,,,,,,"
                        Log_PosInfoLine (CStr(LineCounter) & "=" & PosInfoLine)
                        Call ff.WriteINI(PosInfoFile, "Worklist", CStr(LineCounter), PosInfoLine)
                        LineCounter = LineCounter + 1
                    End If
                End With
            Next a

            Call Check_Empty_TubeSet(CurrentTubeSet)
        Loop
        Call ff.WriteINI(PosInfoFile, "Worklist", "Num", CStr(LineCounter - 1))


OutOfBlue:
    MAXDispensedTarget = OverAllDispensedQuotas Mod MAX_RACK_POS
    MAXDispensedRack = Int(OverAllDispensedQuotas / MAX_RACK_POS)

    Call Check_Empty_TubeSet(CurrentTubeSet)
    Call Log_TubeSet_IsEmpty(CurrentTubeSet)
    Call Empty_Tubeset(CurrentTubeSet)
    db.Close


                '*****************************************************
                '
                ' Dispense Liquid
                '
    'Dim Par1 As Variant, Par2 As Variant, Par3 As Variant
    'Dim FunctName$

    SUBPROCESS_ACTIVE = True

    'FunctName = ""
    'If FunctName="" Then
    '   FunctName="main"
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
        Call XToolApp.CallSub("C:\Lirix\data\process\Distribucion_BB_001.bas", "main")
    'End If
    SUBPROCESS_ACTIVE = False
                '*****************************************************
    n = n + NumberOfTips
    Call xs.DropTips(AllChannels, 0, NO, NO, 0, NO, YES, NO, YES, YES)

Next i
StrXMLFileName = Create_XML_File(strFilePath, "Muestras del " & Format(Now, "dd mmmm yyyy hh_nn_ss") & ".XML", vExecution_ID)

'***********************************************************************************
    '* davide 20/12/2010: proceso de creación de fichero CSV liquido                   *
    '***********************************************************************************
    '*---------------------------------------------------------------------------------*
    '* esta función acepta el nombre del mdb el nombre del fichero csv y el num.       *
    '* de ejecución que por defecto es el último que se ha hecho                       *
    '*---------------------------------------------------------------------------------*
    'START CODING HERE -----------------------------------------------------------------

     Call Create_VITROSOFT_CSV_File(strFilePath, "scan.csv", vExecution_ID)

    '***********************************************************************************
    '* davide 20/12/2010: proceso de creación de fichero CSV liquido                   *
    '***********************************************************************************


Call Execute_HTML_File(StrXMLFileName)

Set db = Nothing
Set Dbs = Nothing
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
        MsgBox ("Oooops, there's something wrong here")
        Err.Raise Err.Number
End Select
        xs.MoveAbsPos (HomePosition)
    Call Cleanup

End Sub


Sub Insert_Record(vExecution_ID As Long, vBarCode As String, vSourceRack As String, vSourceTube As Integer, vTargetRack As String, vTargetRack_ID As String, vWellNumber As Integer, vPosition As String, vCurrentQuota As Long, db As DAO.Database)
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
    rs!TargetRack_ID = vTargetRack_ID
    rs!WellNumber = vWellNumber
    rs!Position = vPosition
    rs!QuotaVolume = vCurrentQuota

    rs.Update

    Set rs = Nothing

End Sub
Sub FillWellMatrix(WellMatrix() As String)

      WellMatrix(1) = "A1"
      WellMatrix(2) = "B1"
      WellMatrix(3) = "C1"
      WellMatrix(4) = "D1"
      WellMatrix(5) = "E1"
      WellMatrix(6) = "F1"
      WellMatrix(7) = "G1"
      WellMatrix(8) = "H1"
      WellMatrix(9) = "A2"
      WellMatrix(10) = "B2"
      WellMatrix(11) = "C2"
      WellMatrix(12) = "D2"
      WellMatrix(13) = "E2"
      WellMatrix(14) = "F2"
      WellMatrix(15) = "G2"
      WellMatrix(16) = "H2"
      WellMatrix(17) = "A3"
      WellMatrix(18) = "B3"
      WellMatrix(19) = "C3"
      WellMatrix(20) = "D3"
      WellMatrix(21) = "E3"
      WellMatrix(22) = "F3"
      WellMatrix(23) = "G3"
      WellMatrix(24) = "H3"
      WellMatrix(25) = "A4"
      WellMatrix(26) = "B4"
      WellMatrix(27) = "C4"
      WellMatrix(28) = "D4"
      WellMatrix(29) = "E4"
      WellMatrix(30) = "F4"
      WellMatrix(31) = "G4"
      WellMatrix(32) = "H4"
      WellMatrix(33) = "A5"
      WellMatrix(34) = "B5"
      WellMatrix(35) = "C5"
      WellMatrix(36) = "D5"
      WellMatrix(37) = "E5"
      WellMatrix(38) = "F5"
      WellMatrix(39) = "G5"
      WellMatrix(40) = "H5"
      WellMatrix(41) = "A6"
      WellMatrix(42) = "B6"
      WellMatrix(43) = "C6"
      WellMatrix(44) = "D6"
      WellMatrix(45) = "E6"
      WellMatrix(46) = "F6"
      WellMatrix(47) = "G6"
      WellMatrix(48) = "H6"
      WellMatrix(49) = "A7"
      WellMatrix(50) = "B7"
      WellMatrix(51) = "C7"
      WellMatrix(52) = "D7"
      WellMatrix(53) = "E7"
      WellMatrix(54) = "F7"
      WellMatrix(55) = "G7"
      WellMatrix(56) = "H7"
      WellMatrix(57) = "A8"
      WellMatrix(58) = "B8"
      WellMatrix(59) = "C8"
      WellMatrix(60) = "D8"
      WellMatrix(61) = "E8"
      WellMatrix(62) = "F8"
      WellMatrix(63) = "G8"
      WellMatrix(64) = "H8"
      WellMatrix(65) = "A9"
      WellMatrix(66) = "B9"
      WellMatrix(67) = "C9"
      WellMatrix(68) = "D9"
      WellMatrix(69) = "E9"
      WellMatrix(70) = "F9"
      WellMatrix(71) = "G9"
      WellMatrix(72) = "H9"
      WellMatrix(73) = "A10"
      WellMatrix(74) = "B10"
      WellMatrix(75) = "C10"
      WellMatrix(76) = "D10"
      WellMatrix(77) = "E10"
      WellMatrix(78) = "F10"
      WellMatrix(79) = "G10"
      WellMatrix(80) = "H10"
      WellMatrix(81) = "A11"
      WellMatrix(82) = "B11"
      WellMatrix(83) = "C11"
      WellMatrix(84) = "D11"
      WellMatrix(85) = "E11"
      WellMatrix(86) = "F11"
      WellMatrix(87) = "G11"
      WellMatrix(88) = "H11"
      WellMatrix(89) = "A12"
      WellMatrix(90) = "B12"
      WellMatrix(91) = "C12"
      WellMatrix(92) = "D12"
      WellMatrix(93) = "E12"
      WellMatrix(94) = "F12"
      WellMatrix(95) = "G12"
      WellMatrix(96) = "H12"

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
    Fld.Attributes = dbAutoIncrField
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
    Fld.Attributes = dbAutoIncrField
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
    Set Fld = Tbl.CreateField("TargetRack_ID", dbText)
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
    Set ind = Tbl.CreateIndex("UK_Distribution_Lines_Position")
    With ind
        .Fields.Append .CreateField("Execution_ID")
        .Fields.Append .CreateField("TargetRack_ID")
        .Fields.Append .CreateField("Position")
        .Unique = True
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
        .Attributes = dbRelationUpdateCascade + dbRelationDeleteCascade

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
    qdf.SQL = "SELECT [Execution_ID], [BarCode], [SourceRack], [SourceTube], [TargetRack], [TargetRack_ID], [Position], [WellNumber], [QuotaVolume] FROM Distribution_Lines  ORDER BY [TargetRack], [WellNumber];"

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
    Call rts.open("SELECT 1 AS [dum my], [TargetRack_ID] AS [Placa destino], [SourceRack] AS Origen, [SourceTube] AS [Posición origen], [BarCode] AS SampleID, [Position] AS [Posición destino], [QuotaVolume] AS Volumen FROM qry_ExportHTML WHERE [Execution_ID]= " & vExecution_ID, conn)
    'Save the Recordset into a DOM tree
    Call rts.Save(xmlDoc, 1)
    Call xmlDoc.insertBefore(xmlDoc.createProcessingInstruction("xml-stylesheet", "type=""text/xsl"" href=""./xsl/INyDIA_Distribute_Log.XSL"""), xmlDoc.documentElement)

    'Writes the datetime of the creation
    Dim xmlFechaNode As IXMLDOMNode
    Set xmlFechaNode = xmlDoc.documentElement.appendChild(xmlDoc.createNode(NODE_ELEMENT, "fecha_hora", ""))
    xmlFechaNode.Text = Format(Now, "dddd dd mmmm - hh:nn")

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
        objFSO.DeleteFile (strFilePath)
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
     Debug.Print "Searched string Not In the source file"
    End If

End Sub

Function GetFile(FileName As String)
Dim FS As Object
     If FileName <> "" Then
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
Const ForReading = 1
Const ForWriting = 2
Const ForAppending = 8

Dim OutStream, FS As Object

     'On Error Resume Next
     Set FS = CreateObject("Scripting.FileSystemObject")
       Set OutStream = FS.OpenTextFile(FileName, ForAppending, True)
       OutStream.WriteLine Contents
       OutStream.Close
     Set OutStream = Nothing
     Set FS = Nothing

End Function
Sub Check_File(strFilePath As String)
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not (objFSO.FileExists(strFilePath)) Then
        Create_MDB_File (strFilePath)
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
    rs!Execution_Notes = "y-Plat Aliq: distribución 1 a N"
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
              "Line #: " & .Line & vbCrLf & _
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
Function Lee_Numero_de_Muestras(fchMuestras As String) As Integer
'Dim ff
Dim str_Result

'Set ff = New cuf.FileFunctions
str_Result = ff.GetINIString(fchMuestras, "MUESTRAS", "NoMuestras")

'    If str_Result = "" Then
'        str_Result = InputBox$("Número de muestras", "Valor no encontrado")
'    End If
    Lee_Numero_de_Muestras = CInt(str_Result)
'Set ff = Nothing

'Dim ff

    'Set ff = CreateObject("cuf.FileFunctions")

    'Lee_Numero_de_Muestras = CInt(ff.GetINIString(fchMuestras, "MUESTRAS", "NoMuestras"))
    'Set ff = Nothing
End Function


Sub Create_VITROSOFT_CSV_File(strFilePath As String, StrCSVFileName As String, vExecution_ID) 'As Boolean
Const adClipString As Integer = 2
Const Separador As String = ";"
Const vbComma As String = ","
Const vbQuotes As String = """"
Const vbDblQuotes As String = """"""

' Const strDummy As String = "ORIGEN_ADN;;013/032403"
' en vez que origen_adn nombre rack nuestro? 013/032403 en blanco
' solo sample id sin guión Sin nada
' en lineas > 23 los campos que no interesa en blanco
' en algun espacio nuestro nombre placa destino


    Dim Datos_Csv As String
    Dim conn As Object, rts As Object, i As Integer, rtsFiltered As Object
    Dim filesys As Object
    Dim varPathCurrent As String
    Dim RackCounter As Integer, TubeCounter As Integer, CuotaCounter As Integer, LineCounter As Integer, PreviousRack As String, PreviousBarcode As String, strVITROLine As String
    Dim RackNums() As String
    Dim RackNames() As String

    RackCounter = 0
    TubeCounter = 0
    CuotaCounter = 0
    LineCounter = 0

    Set filesys = CreateObject("Scripting.FileSystemObject")
    varPathCurrent = filesys.GetParentFolderName(strFilePath)
    Set filesys = Nothing

    Call Delete_File(varPathCurrent & "\..\export\" & StrCSVFileName)
    Set conn = CreateObject("ADODB.Connection")
    conn.Provider = "Microsoft.Jet.OLEDB.4.0"
    Call conn.open(strFilePath)

    Set rts = CreateObject("ADODB.recordset")
    Call rts.open("SELECT [TargetRack], [TargetRack_ID], [Position], [SourceRack], [SourceTube], [BarCode], [QuotaVolume], [WellNumber] FROM qry_ExportHTML WHERE [Execution_ID]= " & vExecution_ID, conn)

    ' Devuelve los datos separados por comas y con un salto de carro
    ' Datos_Csv = rts.GetString(adClipString, -1, ",", vbCrLf, "(NULL)")
    Create_VITROSOFT_CSV_Cabecera (varPathCurrent & "\..\export\scan.csv")

    Call rts.MoveFirst

    Do While Not rts.EOF
        If rts("TargetRack_ID") = PreviousRack Then
            TubeCounter = TubeCounter + 1
        Else
            RackCounter = RackCounter + 1
            TubeCounter = 1
            ReDim Preserve RackNums(RackCounter - 1)
            ReDim Preserve RackNames(RackCounter - 1)
            RackNums(RackCounter - 1) = rts("TargetRack_ID")
            RackNames(RackCounter - 1) = rts("TargetRack")
        End If
        PreviousRack = rts("TargetRack_ID")

        If rts("BarCode") = PreviousBarcode Then
            CuotaCounter = CuotaCounter + 1
        Else
            CuotaCounter = 1
            LineCounter = LineCounter + 1
            strVITROLine = CStr(RackCounter) & Separador & SpaceFill(1) & Separador & SpaceFill(LineCounter) & Separador & rts("SourceRack") & Separador & Separador & Separador & rts("BarCode") ' & "-" & CuotaCounter

            Call WriteFile(varPathCurrent & "\..\export\" & StrCSVFileName, strVITROLine)
        End If
        PreviousBarcode = rts("BarCode")

'¡          strVITROLine = CStr(RackCounter) & Separador & SpaceFill(1) & Separador & SpaceFill(TubeCounter) & Separador & rts("SourceRack") & Separador & Separador & Separador & rts("BarCode") ' & "-" & CuotaCounter
'¡          'Debug.Print strVITROLine
'¡
'¡          WriteFile(varPathCurrent & "\..\export\" & StrCSVFileName, strVITROLine)
        rts.MoveNext
    Loop

    For i = 0 To UBound(RackNums)
            strVITROLine = CStr(23 + i) & Separador & SpaceFill(1) & Separador & SpaceFill(0) & Separador & Separador & RackNames(i) & Separador & Separador & RackNums(i)
            Call WriteFile(varPathCurrent & "\..\export\" & StrCSVFileName, strVITROLine)
            rts.Filter = "TargetRack_ID" & " = '" & RackNums(i) & "'"
            Call Delete_File(varPathCurrent & "\..\export\" & RackNums(i) & ".csv")

            Call rts.MoveFirst
                Do While Not rts.EOF
                strVITROLine = CStr(rts("WellNumber")) & vbComma & vbDblQuotes & vbComma & vbComma & vbQuotes & RackNames(i) & vbQuotes & vbComma & rts("SourceTube") & vbComma & vbQuotes & rts("BarCode") & vbQuotes & vbComma & rts("QuotaVolume") & vbComma & CStr(0) & vbComma & vbQuotes & rts("SourceRack") & vbQuotes & String(5, ",")
                'Debug.Print strVITROLine
                Call WriteFile(varPathCurrent & "\..\export\" & RackNums(i) & ".csv", strVITROLine)
                rts.MoveNext
            Loop

    Next

    Set rts = Nothing
    Call conn.Close
    Set conn = Nothing

Exit Sub
'Error
errFunction:

MsgBox Err.Description, vbCritical

End Sub
Sub Create_VITROSOFT_CSV_Cabecera(StrCSVFileName As String)
    Call WriteFile(StrCSVFileName, "") ' linea en blanco
End Sub

Function SpaceFill(n As Integer) As String
    SpaceFill = String(2 - (Len(CStr(n))), " ") & CStr(n)
End Function
Function Exists_FileINI(strFilePath As String) As Boolean
Dim objFSO As Object, LastModifiedDate As Date
'El fichero INI puede ser muy antiguo no importa

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If (objFSO.FileExists(strFilePath)) Then
'       LastModifiedDate = objFSO.GetFile(strFilePath).DateLastModified
'
'       If DateDiff("h", LastModifiedDate, Now) > 1 Then
'           Dim MsgRes As VbMsgBoxResult
'           MsgRes = MsgBox(strFilePath & " fue modificado hace " & DateDiff("h", LastModifiedDate, Now) & " horas. ¿Continuar?" ,vbOkCancel,"La fecha/hora del fichero es demasiado antigua")
'           If MsgRes=vbCancel Then
'               xs.AbortRunDetected = True
'               Exists_File = False
'           Else
'               Exists_File = True
'           End If
'       Else
            Exists_FileINI = True
'       End If
    Else
        Exists_FileINI = False
    End If
    Set objFSO = Nothing

End Function
Function Lee_Fichero_INI_str(fchINI As String, strVariable As String, strMessage As String) As String
Dim str_Result As String ', ff As cuf.FileFunctions

If (ff Is Nothing) Then
    Set ff = New cuf.FileFunctions
End If
    str_Result = ff.GetINIString(fchINI, "INyDIA", strVariable)

    If str_Result = "" Then
        str_Result = InputBox$(strMessage, "Valor no encontrado")
        If str_Result = "" Then
            Call Err.Raise(-1572, "y-Plat Aliquot", "Valor erróneo u inexistente para " & strVariable)
        End If
    End If

    Debug.Print strVariable & String$(20 - Len(strVariable), "_") & ":" & vbTab & str_Result
    Lee_Fichero_INI_str = str_Result

'Set ff = Nothing
End Function
Function Exists_File(strFilePath As String) As Boolean
Dim objFSO As Object, LastModifiedDate As Date
'Hay que comprobar que el fichero no sea antiguo

    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If (objFSO.FileExists(strFilePath)) Then
        LastModifiedDate = objFSO.GetFile(strFilePath).DateLastModified

        If DateDiff("h", LastModifiedDate, Now) > 1 Then
            Dim MsgRes As VbMsgBoxResult
            MsgRes = MsgBox(strFilePath & " fue modificado hace " & DateDiff("h", LastModifiedDate, Now) & " horas. ¿Continuar?", vbOKCancel, "La fecha/hora del fichero es demasiado antigua")
            If MsgRes = vbCancel Then
                'xs.AbortRunDetected = True
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
Function Lee_Codigos_de_Placas(fchplateBC As String, MP As String) As String
'Dim ff
Dim codPlaca As String

    'Set ff = CreateObject("cuf.FileFunctions")
    codPlaca = ff.GetINIString(fchplateBC, "PLATES BC", MP)
    If codPlaca <> "" Then
        Lee_Codigos_de_Placas = codPlaca
    Else
        Call Err.Raise(-1573, "y-Plat Aliquot", "Valor erróneo u inexistente para código de placa " & MP)
    End If
    'Set ff = Nothing

End Function
Function Lee_Numero_de_Muestras_enbcdata(fchbcdata As String) As Integer
'Dim ff

'   Set ff = CreateObject("cuf.FileFunctions")

    Lee_Numero_de_Muestras_enbcdata = CInt(ff.GetINIString(fchbcdata, "Barcode", "Num"))
'   Set ff = Nothing
End Function
