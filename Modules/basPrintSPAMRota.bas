Attribute VB_Name = "basPrintSPAMRota"
Option Explicit

Dim rstPrintSPAMRota As Recordset, NameFormat As Byte

Public Function CreateRotaPrintTable()
Dim ErrCode As Integer
    On Error GoTo ErrorTrap
    
    Set rstPrintSPAMRota = Nothing
    
    '
    'Recreate Weightings table structure & fields
    '
    If Not DeleteTable("tblPrintSPAMRota") Then
        CreateRotaPrintTable = False
        Exit Function
    End If
    If Not CreateTable(ErrCode, "tblPrintSPAMRota", "RotaDate", "TEXT") Then
        CreateRotaPrintTable = False
        Exit Function
    End If
    If Not CreateField(ErrCode, "tblPrintSPAMRota", "Attendants", "MEMO") Then
        CreateRotaPrintTable = False
        Exit Function
    End If
    If Not CreateField(ErrCode, "tblPrintSPAMRota", "Microphones", "MEMO") Then
        CreateRotaPrintTable = False
        Exit Function
    End If
    If Not CreateField(ErrCode, "tblPrintSPAMRota", "Sound", "MEMO") Then
        CreateRotaPrintTable = False
        Exit Function
    End If
    If Not CreateField(ErrCode, "tblPrintSPAMRota", "Platform", "MEMO") Then
        CreateRotaPrintTable = False
        Exit Function
    End If
    If Not CreateField(ErrCode, "tblPrintSPAMRota", "ID", "LONG") Then
        CreateRotaPrintTable = False
        Exit Function
    End If

    

    Exit Function
ErrorTrap:
    EndProgram



End Function
Public Function CreateRotaPrintTableForBros()
Dim ErrCode As Integer
    On Error GoTo ErrorTrap
    
    Set rstPrintSPAMRota = Nothing
    
    If Not DeleteTable("tblPrintSPAMRota") Then
        CreateRotaPrintTableForBros = False
        Exit Function
    End If
    If Not CreateTable(ErrCode, "tblPrintSPAMRota", "PersonID", "LONG") Then
        CreateRotaPrintTableForBros = False
        Exit Function
    End If
    If Not CreateField(ErrCode, "tblPrintSPAMRota", "BroName", "TEXT") Then
        CreateRotaPrintTableForBros = False
        Exit Function
    End If
    If Not CreateField(ErrCode, "tblPrintSPAMRota", "Attendants", "MEMO") Then
        CreateRotaPrintTableForBros = False
        Exit Function
    End If
    If Not CreateField(ErrCode, "tblPrintSPAMRota", "Microphones", "MEMO") Then
        CreateRotaPrintTableForBros = False
        Exit Function
    End If
    If Not CreateField(ErrCode, "tblPrintSPAMRota", "Platform", "MEMO") Then
        CreateRotaPrintTableForBros = False
        Exit Function
    End If
    If Not CreateField(ErrCode, "tblPrintSPAMRota", "Sound", "MEMO") Then
        CreateRotaPrintTableForBros = False
        Exit Function
    End If

    Exit Function
ErrorTrap:
    EndProgram

End Function

Public Sub PopulateRotaForPrint()
Dim i As Integer, MaxJobs As Byte, TempArray, AttCount As Integer, MicCount As Integer, PltCount As Integer, SndCount As Integer
Dim n As Integer, PrevMic As Integer, PrevSnd As Integer, PrevAtt As Integer, PrevPlt As Integer, PrevDate, j As Integer
Dim k As Long, StoreName As String, FirstJob As Integer, DateFormat As Byte

On Error GoTo ErrorTrap
    
    NameFormat = GlobalParms.GetValue("SPAMRotaNameDisplayFormat", "NumVal")
    Set rstRota = Nothing
    
    Set rstRota = CMSDB.OpenRecordset("SELECT * " & _
                                            "FROM tblRota " _
                                                           , dbOpenDynaset)
    
    Set rstPrintSPAMRota = CMSDB.OpenRecordset("SELECT * " & _
                                                            "FROM tblPrintSPAMRota " _
                                                                                , dbOpenDynaset)
    PrevMic = 0
    PrevSnd = 0
    PrevPlt = 0
    PrevAtt = 0
    PrevDate = ""
    
    AcquireRotaStructure AttCount, MicCount, SndCount, PltCount, FirstJob
    
    DateFormat = GlobalParms.GetValue("SPAMDateFormat", "NumVal")
    
    lnkNameFormatForPrint = NameFormat
    lnkDateFormatForPrint = DateFormat
    
    TempArray = Array(AttCount, MicCount, PltCount, SndCount)
    
    MaxJobs = FindMaxVal(TempArray)
    
    rstRota.MoveLast
    rstRota.MoveFirst
    
    '
    'Put in dates
    '
    With rstPrintSPAMRota
    For n = 0 To rstRota.RecordCount - 1
        For i = 1 To MaxJobs
            .AddNew
            If i = 1 Then 'Put in date only for 1st line of each group
                Select Case DateFormat
                Case 1:
                    !RotaDate = Format$(rstRota!RotaDate, "MMM DD")
                Case 2:
                    !RotaDate = Format$(rstRota!RotaDate, "ddd MMM DD")
                End Select
            Else
                !RotaDate = ""
            End If
            !Attendants = ""
            !Microphones = ""
            !Platform = ""
            !Sound = ""
            !ID = n
            .Update
        Next i
        rstRota.MoveNext
    Next n
    
    CMSDB.TableDefs.Refresh
    
    rstRota.MoveFirst
    .MoveFirst
    i = 0
    k = 0
    
    For n = 0 To rstRota.RecordCount - 1
        For j = 2 To 5 'For each column of tblPrintSPAMRota....
            Do
                i = i + 1
                .Edit
                Select Case j
                Case 2:   'Column 2 (Attendants)
                    If i <= AttCount Then
                        If PrevAtt <> rstRota.Fields("Attendant_" & Format$(i, "00")) Then
                            !Attendants = GetName(rstRota.Fields("Attendant_" & Format$(i, "00")))
                            PrevAtt = rstRota.Fields("Attendant_" & Format$(i, "00"))
                            .Update
                            If Left(!Attendants, 1) = "?" Then
                                SPAMBroNotFound = True
                            End If
                            PrevDate = !RotaDate
                            .MoveNext
                        End If
                    End If
                
                Case 3:   'Mics
                    If i <= MicCount Then
                        If PrevMic <> rstRota.Fields("RovingMic_" & Format$(i, "00")) Then
                            !Microphones = GetName(rstRota.Fields("RovingMic_" & Format$(i, "00")))
                            PrevMic = rstRota.Fields("RovingMic_" & Format$(i, "00"))
                            .Update
                            If Left(!Microphones, 1) = "?" Then
                                SPAMBroNotFound = True
                            End If
                            PrevDate = !RotaDate
                            .MoveNext
                        End If
                    End If
                    
                Case 4:   'Platform
                    If i <= PltCount Then
                        If PrevPlt <> rstRota.Fields("Platform_" & Format$(i, "00")) Then
                            !Platform = GetName(rstRota.Fields("Platform_" & Format$(i, "00")))
                            PrevPlt = rstRota.Fields("Platform_" & Format$(i, "00"))
                            .Update
                            If Left(!Platform, 1) = "?" Then
                                SPAMBroNotFound = True
                            End If
                            PrevDate = !RotaDate
                            .MoveNext
                        End If
                    End If
                    
                Case 5:    'sound
                    If i <= SndCount Then
                        If PrevSnd <> rstRota.Fields("Sound_" & Format$(i, "00")) Then
                            !Sound = GetName(rstRota.Fields("Sound_" & Format$(i, "00")))
                            PrevSnd = rstRota.Fields("Sound_" & Format$(i, "00"))
                            .Update
                            If Left(!Sound, 1) = "?" Then
                                SPAMBroNotFound = True
                            End If
                            PrevDate = !RotaDate
                            .MoveNext
                        End If
                    End If
                End Select
                
            Loop Until i = MaxJobs
            i = 0
            .FindFirst "ID = " & rstRota!SeqNum - 1
        Next j
        
            
        
        .FindFirst "ID = " & rstRota!SeqNum
        
        If .NoMatch Then
            Exit For
        End If
        
        rstRota.MoveNext
        
        PrevMic = 0
        PrevSnd = 0
        PrevPlt = 0
        PrevAtt = 0

    Next n
        
    End With
            

    Exit Sub
ErrorTrap:
    EndProgram

End Sub
Public Sub PopulateRotaForPrint2()
Dim i As Integer, MaxJobs As Byte, TempArray, AttCount As Integer, MicCount As Integer, PltCount As Integer, SndCount As Integer
Dim n As Integer, PrevMic As Integer, PrevSnd As Integer, PrevAtt As Integer, PrevPlt As Integer, PrevDate, j As Integer
Dim k As Long, StoreName As String, FirstJob As Integer, DateFormat As Byte
Dim sAtts As String, sMics As String, sPlat As String, sSnd As String, str As String

On Error GoTo ErrorTrap
    
    NameFormat = GlobalParms.GetValue("SPAMRotaNameDisplayFormat", "NumVal")
    Set rstRota = Nothing
    
    Set rstRota = CMSDB.OpenRecordset("SELECT * " & _
                                            "FROM tblRota " _
                                                           , dbOpenDynaset)
    
    Set rstPrintSPAMRota = CMSDB.OpenRecordset("SELECT * " & _
                                                            "FROM tblPrintSPAMRota " _
                                                                                , dbOpenDynaset)
    PrevMic = 0
    PrevSnd = 0
    PrevPlt = 0
    PrevAtt = 0
    PrevDate = ""
    
    AcquireRotaStructure AttCount, MicCount, SndCount, PltCount, FirstJob
    
    DateFormat = GlobalParms.GetValue("SPAMDateFormat", "NumVal")
    
    lnkNameFormatForPrint = NameFormat
    lnkDateFormatForPrint = DateFormat
    
    TempArray = Array(AttCount, MicCount, PltCount, SndCount)
    
    MaxJobs = FindMaxVal(TempArray)
    
    rstRota.MoveLast
    rstRota.MoveFirst
    
    '
    'Put in dates
    '
    With rstPrintSPAMRota
    For n = 0 To rstRota.RecordCount - 1
        .AddNew
        Select Case DateFormat
        Case 1:
            !RotaDate = Format$(rstRota!RotaDate, "MMM DD")
        Case 2:
            !RotaDate = Format$(rstRota!RotaDate, "ddd MMM DD")
        End Select
        !Attendants = ""
        !Microphones = ""
        !Platform = ""
        !Sound = ""
        !ID = n
        .Update
        rstRota.MoveNext
    Next n
    
    CMSDB.TableDefs.Refresh
    
    rstRota.MoveFirst
    .MoveFirst
    i = 0
    k = 0
    
    For n = 0 To rstRota.RecordCount - 1
        .Edit
        For j = 2 To 5 'For each column of tblPrintSPAMRota....
            Select Case j
            Case 2:   'Column 2 (Attendants)
                For i = 2 To 1 + AttCount
                    str = GetName(rstRota.Fields(i))
                    If Left(str, 1) = "?" Then
                        SPAMBroNotFound = True
                    End If
                    sAtts = sAtts & str
                    If i < 1 + AttCount Then
                        sAtts = sAtts & vbCrLf
                    End If
                Next i
                !Attendants = sAtts
            
            Case 3:   'Mics
                For i = 2 + AttCount To 1 + AttCount + MicCount
                    str = GetName(rstRota.Fields(i))
                    If Left(str, 1) = "?" Then
                        SPAMBroNotFound = True
                    End If
                    sMics = sMics & str
                    If i < 1 + AttCount + MicCount Then
                        sMics = sMics & vbCrLf
                    End If
                Next i
                !Microphones = sMics
                
            Case 4:   'sound
                For i = 2 + AttCount + MicCount To 1 + AttCount + MicCount + SndCount
                    str = GetName(rstRota.Fields(i))
                    If Left(str, 1) = "?" Then
                        SPAMBroNotFound = True
                    End If
                    sSnd = sSnd & str
                    If i < 1 + AttCount + MicCount + SndCount Then
                        sSnd = sSnd & vbCrLf
                    End If
                Next i
                !Sound = sSnd
                
            Case 5:    'Platform
                For i = 2 + AttCount + MicCount + SndCount To 1 + AttCount + MicCount + SndCount + PltCount
                    str = GetName(rstRota.Fields(i))
                    If Left(str, 1) = "?" Then
                        SPAMBroNotFound = True
                    End If
                    sPlat = sPlat & str
                    If i < 1 + AttCount + MicCount + PltCount + SndCount Then
                        sPlat = sPlat & vbCrLf
                    End If
                Next i
                !Platform = sPlat
            End Select
        Next j
        .Update
        i = 0
        sSnd = ""
        sPlat = ""
        sMics = ""
        sAtts = ""
        .MoveNext
        rstRota.MoveNext
        
    Next n
        
    End With
            

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Public Sub PopulateRotaForPrintForBros()
Dim i As Integer, MaxJobs As Byte, TempArray, AttCount As Integer, MicCount As Integer, PltCount As Integer, SndCount As Integer
Dim n As Integer, j As Integer
Dim k As Long, StoreName As String, FirstJob As Integer, DateFormat As Byte
Dim rsBroList As Recordset, b1stRecDone As Boolean

On Error GoTo ErrorTrap
    
    NameFormat = GlobalParms.GetValue("SPAMRotaNameDisplayFormat", "NumVal")
    Set rstRota = Nothing
    
    Set rstRota = CMSDB.OpenRecordset("SELECT * " & _
                                            "FROM tblRota " _
                                                           , dbOpenDynaset)
    
    Set rstPrintSPAMRota = CMSDB.OpenRecordset("SELECT * " & _
                                                            "FROM tblPrintSPAMRota " _
                                                                                , dbOpenDynaset)
                                                                                
    Set rsBroList = GetGeneralRecordset("SELECT DISTINCT Person " & _
                                        "FROM (" & _
                                        "SELECT Person FROM tblTaskAndPerson a " & _
                                        "INNER JOIN tblNameAddress b " & _
                                        " ON a.Person = b.ID " & _
                                        "WHERE Task IN (57,58,59,60) " & _
                                        "ORDER BY LastName, FirstName, MiddleName) ")
                                        
    If rsBroList.BOF Then
        ShowMessage "No brothers set up", 1500, frmSoundAndPlatformRota
        GoTo ClearOff
    End If
    If rstRota.BOF Then
        ShowMessage "Rota is empty", 1500, frmSoundAndPlatformRota
        GoTo ClearOff
    End If
                                            
                                            
    AcquireRotaStructure AttCount, MicCount, SndCount, PltCount, FirstJob
    
    DateFormat = GlobalParms.GetValue("SPAMDateFormat", "NumVal")
    
    lnkNameFormatForPrint = NameFormat
    lnkDateFormatForPrint = DateFormat
    
    TempArray = Array(AttCount, MicCount, PltCount, SndCount)
    
    MaxJobs = FindMaxVal(TempArray)
        
    '
    'Put in names
    '
    With rstPrintSPAMRota
    Do Until rsBroList.EOF
        rstPrintSPAMRota.AddNew
        rstPrintSPAMRota!PersonID = rsBroList!Person
        rstPrintSPAMRota!BroName = CongregationMember.NameWithMiddleInitial(rsBroList!Person)
        rstPrintSPAMRota!Attendants = ""
        rstPrintSPAMRota!Microphones = ""
        rstPrintSPAMRota!Sound = ""
        rstPrintSPAMRota!Platform = ""
        rstPrintSPAMRota.Update
        rsBroList.MoveNext
        b1stRecDone = True
    Loop
    
    rstRota.MoveFirst
    .MoveFirst
    i = 0
    k = 0
    
    Do Until rstRota.EOF
        For j = 3 To 6 'For each column of tblPrintSPAMRota....
            Do
                i = i + 1
                Select Case j
                Case 3:   'Column 3 (Attendants)
                    If i <= AttCount Then
                        .FindFirst "PersonID = " & rstRota.Fields("Attendant_" & Format$(i, "00"))
                        If Not .NoMatch Then
                            .Edit
                            !Attendants = AddStringListDelimiter(!Attendants, rstRota!RotaDate)
                            .Update
                        End If
                    End If
                
                Case 4:   'Mics
                    If i <= MicCount Then
                        .FindFirst "PersonID = " & rstRota.Fields("RovingMic_" & Format$(i, "00"))
                        If Not .NoMatch Then
                            .Edit
                            !Microphones = AddStringListDelimiter(!Microphones, rstRota!RotaDate)
                            .Update
                        End If
                    End If
                    
                Case 5:   'Platform
                    If i <= PltCount Then
                        .FindFirst "PersonID = " & rstRota.Fields("Platform_" & Format$(i, "00"))
                        If Not .NoMatch Then
                            .Edit
                            !Platform = AddStringListDelimiter(!Platform, rstRota!RotaDate)
                            .Update
                        End If
                    End If
                                    
                Case 6:    'sound
                    If i <= SndCount Then
                        .FindFirst "PersonID = " & rstRota.Fields("Sound_" & Format$(i, "00"))
                        If Not .NoMatch Then
                            .Edit
                            !Sound = AddStringListDelimiter(!Sound, rstRota!RotaDate)
                            .Update
                        End If
                    End If
                
                End Select
                
            Loop Until i = MaxJobs
            i = 0
        Next j
                
        rstRota.MoveNext
        
    Loop
        
    End With
    
ClearOff:
    On Error Resume Next
    rstRota.Close
    Set rstRota = Nothing
    rstPrintSPAMRota.Close
    Set rstPrintSPAMRota = Nothing
    rsBroList.Close
    Set rsBroList = Nothing
    
    CMSDB.Execute "DELETE FROM tblPrintSPAMRota " & _
                  "WHERE Attendants = '' " & _
                  "AND Microphones = '' " & _
                  "AND Platform = ''" & _
                  "AND Sound = ''"

    Exit Sub
ErrorTrap:
    EndProgram

End Sub


Public Function GetName(BroId As Long) As String


On Error GoTo ErrorTrap

    Select Case NameFormat
    Case 1
        GetName = CongregationMember.FullName(BroId)
    Case 2
        GetName = CongregationMember.FirstAndLastName(BroId)
    Case 3
        GetName = CongregationMember.NameWithMiddleInitial(BroId)
    Case 4
        GetName = CongregationMember.NameWithTwoFirstInitials(BroId)
    Case 5
        GetName = CongregationMember.NameWithOneFirstInitial(BroId)
    Case 6
        GetName = CongregationMember.PersonsInitials(BroId)
    End Select

    Exit Function
ErrorTrap:
    EndProgram

End Function


