Attribute VB_Name = "SourceTube"
'Const PosInfoFile = "C:\Lirix\data\piplist\" + Format(Now, "yyyymmdd-hhnnss") + ".dat"
Type SourceTubeSet
    SourceRackName As String
    Tubes(7) As INyDIA_SourceTube
    IsEmpty As Boolean
    TotQuotas As Integer
    RemainingQuotas As Integer
End Type

'Sub Main
'   Call Test
'End Sub


Sub Test()

    Dim CurrentTubeSet As SourceTubeSet
    Dim ff As New xuf.FileFunctions
    Dim PosInfoLine As String, DestRackName As String, PreviousPos As Integer, CurrentPos As Integer, LineCounter As Long

    Call Initialize_TubeSet(CurrentTubeSet)
    CurrentTubeSet.SourceRackName = "SampleRack_001"

    Call Generate_RNDTubeset(CurrentTubeSet, True, 500)
    Call Log_TubeSet_Info(CurrentTubeSet)

    Call Check_Empty_TubeSet(CurrentTubeSet)
    Call Log_TubeSet_IsEmpty(CurrentTubeSet)




    Call Set_UseIncompleteQuota(CurrentTubeSet, False)
    Call Log_TubeSet_Info(CurrentTubeSet)

    Call Check_Empty_TubeSet(CurrentTubeSet)
    Call Log_TubeSet_IsEmpty(CurrentTubeSet)


    Call Empty_Tubeset(CurrentTubeSet)
    Call Log_TubeSet_Info(CurrentTubeSet)

    Call Check_Empty_TubeSet(CurrentTubeSet)
    Call Log_TubeSet_IsEmpty(CurrentTubeSet)

End Sub


Sub Check_Empty_TubeSet(CurrentTubeSet As SourceTubeSet)

    Dim a As Integer
    With CurrentTubeSet
        .IsEmpty = False
        For a = 0 To 7
            If Not .Tubes(a).IsTubeEmpty Then Exit Sub
        Next a
        .IsEmpty = True
    End With

End Sub

Sub Calculate_Previous_Tube_Quota(CurrentTubeSet As SourceTubeSet)

    Dim Prevquotas As Integer, a As Integer, b As Integer
    Prevquotas = 0

    For a = 0 To 7
        For b = 0 To a - 1
            Prevquotas = Prevquotas + CurrentTubeSet.Tubes(b).NoOfQuotas
        Next b
        CurrentTubeSet.Tubes(a).PreviousTubesQuotas = Prevquotas
        Prevquotas = 0
    Next a

End Sub

Sub Calculate_TubeSetTotQuotas(CurrentTubeSet As SourceTubeSet)

    Dim a As Integer
    With CurrentTubeSet
        .TotQuotas = 0
        For a = 0 To 7
                .TotQuotas = .TotQuotas + .Tubes(a).NoOfQuotas
        Next a
    End With
End Sub

Sub Initialize_TubeSet(CurrentTubeSet As SourceTubeSet)

    Dim a As Integer
    For a = 0 To 7
        Set CurrentTubeSet.Tubes(a) = New INyDIA_SourceTube
        CurrentTubeSet.Tubes(a).TubeOrderNo = a + 1
    Next a
End Sub

Sub Generate_RNDTubeset(CurrentTubeSet As SourceTubeSet, IncompleteQuota As Boolean, Quota As Long)

    Dim a As Integer
    For a = 0 To 3
        With CurrentTubeSet.Tubes(a)
            .DetectedVolume = CLng(Rnd() * &H7000)
            .ReqQuotaVolume = Quota
            .UseIncompleteQuota = IncompleteQuota
        End With
    Next a
    For a = 4 To 7
        With CurrentTubeSet.Tubes(a)
            .DetectedVolume = CLng(0)
            .ReqQuotaVolume = CLng(500)
            .UseIncompleteQuota = IncompleteQuota
        End With
    Next a
End Sub

Sub Empty_Tubeset(CurrentTubeSet As SourceTubeSet)

    Dim a As Integer
    For a = 0 To 7
            With CurrentTubeSet.Tubes(a)
                .DetectedVolume = CLng(0)
                .ReqQuotaVolume = CLng(500)
                .PreviousTubesQuotas = 0
                .NoOfDispensedQuotas = 0
                .TubeOrderNo = 0
            Debug.Print .NoOfQuotas
            End With
    Next a
    Call Calculate_Previous_Tube_Quota(CurrentTubeSet)
    Call Calculate_TubeSetTotQuotas(CurrentTubeSet)

End Sub

Sub Set_UseIncompleteQuota(CurrentTubeSet As SourceTubeSet, IncompleteQuota As Boolean)

    Dim a As Integer
    For a = 0 To 7
        With CurrentTubeSet.Tubes(a)
            .UseIncompleteQuota = IncompleteQuota
        End With
    Next a
    Call Calculate_Previous_Tube_Quota(CurrentTubeSet)
    Call Calculate_TubeSetTotQuotas(CurrentTubeSet)
End Sub

Sub Log_TubeSet_Info(CurrentTubeSet As SourceTubeSet)

    Dim CurrentLogFile
    Dim a As Integer
    Set CurrentLogFile = New INyDIA_LogFile
    For a = 0 To 7
        With CurrentTubeSet.Tubes(a)
            CurrentLogFile.Write_INFO ( _
                    "Detected Volume = " & _
                    .DetectedVolume & " : " & _
                    "Number Of Quotas " & _
                    .NoOfQuotas & " : (" & .NoOfQuotas + .UseIncompleteQuota & ")" & .ReqQuotaVolume & " ul quotas + " & _
                    .LastQuotaVolume & " ul")
            CurrentLogFile.Write_INFO ( _
                    "Is empty = " & .IsTubeEmpty & ", " & _
                    "Use incomplete quota = " & .UseIncompleteQuota)
            CurrentLogFile.Write_INFO ( _
            "Previous Tubes Quotas = " & .PreviousTubesQuotas)
        End With
    Next a
    Set CurrentLogFile = Nothing
End Sub

Sub Log_TubeSet_IsEmpty(CurrentTubeSet As SourceTubeSet)
    Dim CurrentLogFile
    Set CurrentLogFile = New INyDIA_LogFile
        CurrentLogFile.Write_INFO ("Total Quotas in Tube Set: " & CurrentTubeSet.TotQuotas)
        CurrentLogFile.Write_INFO ("Empty Tube Set: " & CurrentTubeSet.IsEmpty)
    Set CurrentLogFile = Nothing
End Sub

Sub Log_PosInfoLine(PosInfoLine As String)

    Dim CurrentLogFile
    Set CurrentLogFile = New INyDIA_LogFile
    CurrentLogFile.Write_INFO (PosInfoLine)

End Sub

