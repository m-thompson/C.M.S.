Attribute VB_Name = "basEditRota"
Option Explicit

Public IDWeightingsArray() As Integer
Public RotaArray() As Variant, TheBroID As Integer, ASPAMBrother As clsCongregationMember
Public SPAMQuery As String, rstRotaForEdit As Recordset, SPAMBroNotFound As Boolean
Public done As Boolean



Public Sub BuildWeightingTables()
Dim i As Integer

On Error GoTo ErrorTrap

    
    CreateWeightingsTable
    CalcInitWtngs
    

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Public Sub FindSelectedWeeksWeightings(weekno As Integer)
Dim i As Integer, j As Integer


On Error GoTo ErrorTrap

    With rstIDWeight
        
    '
    'Reset all weightings
    '
    CMSDB.Execute ("UPDATE tblIDWeightings " & _
                       "SET Weighting = " & 0)
    
    
    '
    'Move back from current week to start of rota, building weightings for each bro
    '
    For i = weekno To 0 Step -1
        For j = 0 To UBound(RotaArray, 2)    'ie for each column of array rota
            If j > 0 Then
                If RotaArray(i, j - 1) <> RotaArray(i, j) Then
                    UpdateWeighting i, j, weekno
                End If
            Else
                UpdateWeighting i, j, weekno
            End If
        Next j
    Next i
    
    '
    'Move forward from current week to end of rota, building weightings for each bro
    '
    For i = weekno + 1 To UBound(RotaArray, 1)
        For j = 0 To UBound(RotaArray, 2) 'ie for each column of array rota
            If j > 0 Then
                If RotaArray(i, j - 1) <> RotaArray(i, j) Then
                    UpdateWeighting i, j, weekno
                End If
            Else
                UpdateWeighting i, j, weekno
            End If
        Next j
    Next i
    
    End With
    

    Exit Sub
ErrorTrap:
    EndProgram


    
End Sub

Public Sub GenerateListOfPossibleBrothersForJob(JobID As String, SeqNum As Integer, rstSPAMRota As Recordset)
Dim JobQuery As String, SelectedRotaDate, i As Long
    

On Error GoTo ErrorTrap

    '
    'What's the Rota Date? Get it via the SeqNum (Col 5 of grid)
    '
    With rstSPAMRota
    .FindFirst ("SeqNum = " & SeqNum + 1)
    SelectedRotaDate = Format(!RotaDate, "mm/dd/yy")
    End With
    
    '
    'Construct SQL to narrow Bro-list to specific job
    '
    Select Case JobID
    Case "A"
        JobQuery = "WHERE CongNo = " & giGlobalDefaultCong & _
                    " AND TaskCategory = " & 6 & _
                    " AND TaskSubCategory = " & 11 & _
                    " AND Task = " & 60
    Case "M"
        JobQuery = "WHERE CongNo = " & giGlobalDefaultCong & _
                    " AND TaskCategory = " & 6 & _
                    " AND TaskSubCategory = " & 10 & _
                    " AND Task = " & 59
    Case "S"
        JobQuery = "WHERE CongNo = " & giGlobalDefaultCong & _
                    " AND TaskCategory = " & 6 & _
                    " AND TaskSubCategory = " & 10 & _
                    " AND Task = " & 57
    Case "P"
        JobQuery = "WHERE CongNo = " & giGlobalDefaultCong & _
                    " AND TaskCategory = " & 6 & _
                    " AND TaskSubCategory = " & 10 & _
                    " AND Task = " & 58
    End Select

    '
    'Take into account Suspend dates...
    '
    SPAMQuery = "SELECT DISTINCTROW tblIDWeightings.ID, tblNameAddress.FirstName & ' ' & " & _
                         "tblNameAddress.MiddleName &  ' ' & " & _
                         "tblNameAddress.LastName " & _
                "FROM (tblIDWeightings " & _
                "INNER JOIN tblNameAddress ON " & _
                "(tblNameAddress.ID = tblIDWeightings.ID)) " & _
                "INNER JOIN tblTaskPersonSuspendDates ON " & _
                "(tblNameAddress.ID = tblTaskPersonSuspendDates.Person) " & _
                JobQuery & _
                " AND ((IsNull(SuspendStartDate) AND IsNull(SuspendEndDate)) " & _
                "OR (IsNull(SuspendStartDate) AND (SuspendEndDate <= #" & SelectedRotaDate & "#)) " & _
                "OR (IsNull(SuspendEndDate) AND (SuspendStartDate > #" & SelectedRotaDate & "#)) " & _
                "OR (SuspendStartDate > #" & SelectedRotaDate & "# OR SuspendEndDate <= #" & SelectedRotaDate & "#)) " & _
                "ORDER BY LastName"
                
    
    '
    'Show the Bro selection list.
    '
    With frmEditRota
    !cmbSelectNewBro.Visible = True
    HandleListBox.PopulateListBox !cmbSelectNewBro, SPAMQuery, CMSDB, 0, "", False, 1, , , , , , True
    HandleListBox.AutoSizeDropDownWidth !cmbSelectNewBro
    !cmbSelectNewBro.Left = !grdEditRota.CellLeft
    !cmbSelectNewBro.Top = !grdEditRota.CellTop
    !cmbSelectNewBro.Width = !grdEditRota.CellWidth
    !cmbSelectNewBro.text = !grdEditRota.text
    End With
    

    Exit Sub
ErrorTrap:
    EndProgram

    
End Sub

Private Sub UpdateWeighting(i, j, weekno)

On Error GoTo ErrorTrap
    With rstIDWeight
    .FindFirst "ID = " & RotaArray(i, j)
    .Edit
    !Weighting = !Weighting + (GetWkWt(i - weekno) + 1) * 10 * ((!ResponsibilityWeighting + 1) * (!Personal_Wtg + 1))
    .Update
    End With


    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Public Sub SetCellShading(TheGrid, RowsToGroup As Byte, StartRow As Integer, EndRow As Integer, StartCol As Byte, Endcol As Byte, Colour1 As Long, Colour2 As Long)
Dim i As Integer, j As Integer, k As Byte, m As Integer, Colour As Long
'
'This procedure shades horizontal bands of an MSFlexGrid in alternate colours (Colour1 and Colour2). The bands have thickness of 'RowsToGroup'.
' You can set values for start/end column and row. Pass the flexgrid in as a variant.
'

On Error GoTo ErrorTrap

    k = 0
    
    With TheGrid
    
    For i = StartRow To (EndRow - RowsToGroup + 1) Step RowsToGroup
        '
        'flip the colour each time
        '
        Select Case k
        Case 0
            Colour = Colour1
            k = 1
        Case 1
            Colour = Colour2
            k = 0
        End Select
                    
        For m = 0 To RowsToGroup - 1
            For j = StartCol To Endcol
                .Row = i + m
                .col = j
                .CellBackColor = Colour
            Next j
        Next m
    Next i
    
    End With
    

    Exit Sub
ErrorTrap:
    EndProgram

    
    
    
End Sub

Public Sub AcquireRotaStructure2(TempNoAttending As Integer, _
                                 TempNoOnMics As Integer, _
                                 TempNoOnSound As Integer, _
                                 TempNoOnPlatform As Integer, _
                                 FirstJobPos As Integer, _
                                 RotaResultSet)
Dim i As Integer, FirstJobFound As Boolean

On Error GoTo ErrorTrap

    
'
'Structure of current tblRota may be different to that set on tblConstants (eg NoOnMics etc). Therefore,
' acquire the ACTUAL structure from tblRota itself. These values are then used later in the program.
'
    
    FirstJobFound = False
    
    For i = 0 To RotaResultSet.Fields.Count - 1
        If InStr(1, RotaResultSet.Fields(i).Name, "Attendant") Then
            TempNoAttending = TempNoAttending + 1
            FirstJobFound = True
        ElseIf InStr(1, RotaResultSet.Fields(i).Name, "RovingMic") Then
            TempNoOnMics = TempNoOnMics + 1
            FirstJobFound = True
        ElseIf InStr(1, RotaResultSet.Fields(i).Name, "Platform") Then
            TempNoOnPlatform = TempNoOnPlatform + 1
            FirstJobFound = True
        ElseIf InStr(1, RotaResultSet.Fields(i).Name, "Sound") Then
            TempNoOnSound = TempNoOnSound + 1
            FirstJobFound = True
        End If
        If FirstJobFound = False Then
            FirstJobPos = i
        End If
    Next i
        
    If i > 0 Then
        FirstJobPos = FirstJobPos + 1
    End If
    

    Exit Sub
ErrorTrap:
    EndProgram

    End Sub


