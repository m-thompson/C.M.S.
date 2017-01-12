Attribute VB_Name = "basTMSCode2"
Option Explicit
Dim mbExcelWasOpen As Boolean
Dim oExcelApp As Object
Dim oExcelDoc As Object
Dim oExcelNamesSheet As Object
Dim oExcelWkSht As Object


Public Function TMS_UpdateSQDescriptionsFromXLS(bDeleteAllFirst As Boolean, Optional bErrorIfXLSNotThere As Boolean = False) As Boolean

'called from UpgradeDB modules

Dim sPathToXLS As String
Dim iRow As Long
Dim bEnd As Boolean
Dim lSQ As Long
Dim lSQSubPt As Long
Dim sDesc As String
Dim sUpdType As String

On Error GoTo ErrorTrap

    sPathToXLS = SpecialFolder(SpecialFolder_AppData) & "\Congregation Management System\TMS SQ Desc Update.xls"
    
    WriteToLogFile "Update TMS Speech Qualities..."
    
    WriteToLogFile "Path to spreadsheet: " & sPathToXLS

    If Not gFSO.FileExists(sPathToXLS) Then
        WriteToLogFile "Spreadsheet not present"
        If bErrorIfXLSNotThere Then
            Err.Raise vbObjectError + 246, "basTMSCode2.TMS_UpdateSQDescriptionsFromXLS", _
                       "'" & sPathToXLS & "' doesn't exist"
        Else
            TMS_UpdateSQDescriptionsFromXLS = True
            Exit Function
        End If
                
    End If
    
    
    If Not OpenExcelForTMSUpdate Then
        Err.Raise vbObjectError + 247, "basTMSCode2.TMS_UpdateSQDescriptionsFromXLS", _
                   "Error opening Excel"
    End If
    

    Screen.MousePointer = vbHourglass
    
    
    'open the spreadsheet
    On Error Resume Next
    oExcelApp.Visible = False
    Set oExcelDoc = oExcelApp.Application.Workbooks.Open(sPathToXLS)
    If Err.number <> 0 Then
        Err.Raise vbObjectError + 248, "basTMSCode2.TMS_UpdateSQDescriptionsFromXLS", _
                   "Error opening Excel workbook"
    End If
    On Error GoTo ErrorTrap
    
    'check that doc relates to CMS
    If oExcelDoc.Worksheets.Count <> 2 Then
        Err.Raise vbObjectError + 249, "basTMSCode2.TMS_UpdateSQDescriptionsFromXLS", _
                   "Invalid spreadsheet - should be two tabs"
    End If
    
    Set oExcelNamesSheet = oExcelDoc.Worksheets(1)
    
    If oExcelNamesSheet.Name <> "TMS SQ Update" Then
        Err.Raise vbObjectError + 250, "basTMSCode2.TMS_UpdateSQDescriptionsFromXLS", _
                   "Invalid spreadsheet - first tab has incorrect name"
    End If
    
    If oExcelNamesSheet.Range("A1").value <> "CounselPoint" Or _
        oExcelNamesSheet.Range("B1").value <> "CounselSubPoint" Or _
        oExcelNamesSheet.Range("C1").value <> "SubPointDescription" Then
        Err.Raise vbObjectError + 251, "basTMSCode2.TMS_UpdateSQDescriptionsFromXLS", _
                   "Invalid spreadsheet"
    End If
    
    If bDeleteAllFirst Then
        CMSDB.Execute "DELETE FROM tblTMSCounselPointComponents "
    End If
    
    bEnd = False
    iRow = 2
    
    Do Until bEnd
    
        If Not CellEmptyTMS(oExcelNamesSheet.Range("A" & iRow)) Then
        
            lSQ = oExcelNamesSheet.Range("A" & iRow).value
            lSQSubPt = oExcelNamesSheet.Range("B" & iRow).value
            sDesc = oExcelNamesSheet.Range("C" & iRow).value
            sUpdType = oExcelNamesSheet.Range("D" & iRow).value
            
            If lSQ < 1 Or lSQ > 53 Then
                Err.Raise vbObjectError + 254, "basTMSCode2.TMS_UpdateSQDescriptionsFromXLS", _
                           "Invalid SQ number (should be 1-53)"
            End If
            
            If lSQSubPt < 1 Or lSQSubPt > 5 Then
                Err.Raise vbObjectError + 255, "basTMSCode2.TMS_UpdateSQDescriptionsFromXLS", _
                           "Invalid SQ sub-point number (should be 1-5)"
            End If
            
            Select Case sUpdType
            Case "U"
                CMSDB.Execute "UPDATE tblTMSCounselPointComponents " & _
                             "SET SubPointDescription = '" & DoubleUpSingleQuotes(sDesc) & "' " & _
                             "WHERE CounselPoint = " & lSQ & " AND CounselSubPoint = " & lSQSubPt
            Case "I"
                CMSDB.Execute "INSERT INTO tblTMSCounselPointComponents " & _
                              " VALUES (" & lSQ & ", " & lSQSubPt & ", '" & DoubleUpSingleQuotes(sDesc) & "') "
            Case "D"
                CMSDB.Execute "DELETE FROM tblTMSCounselPointComponents " & _
                             "WHERE CounselPoint = " & lSQ & " AND CounselSubPoint = " & lSQSubPt
            Case "L"
                'Leave alone
            Case Else
                Err.Raise vbObjectError + 256, "basTMSCode2.TMS_UpdateSQDescriptionsFromXLS", _
                           "Invalid CMS Action Code (U/I/D)"
            
            End Select
                         
            iRow = iRow + 1

        Else
        
            bEnd = True
            Exit Do
            
        End If
    
    Loop
    
    On Error Resume Next
    
    
    Screen.MousePointer = vbNormal
    
    TMS_UpdateSQDescriptionsFromXLS = True
    
    oExcelDoc.Close
    
    If Not mbExcelWasOpen Then
        oExcelApp.DisplayAlerts = False
        oExcelApp.Quit
    Else
        oExcelApp.Visible = True
    End If
    
    Set oExcelApp = Nothing
    Set oExcelDoc = Nothing
    Set oExcelNamesSheet = Nothing
    Set oExcelWkSht = Nothing

    If gFSO.FileExists(sPathToXLS) Then
        gFSO.DeleteFile sPathToXLS, True
        WriteToLogFile sPathToXLS & " deleted"
    End If


    Exit Function
    
ErrorTrap:
    
    WriteToLogFile ("Error updating SQ sub-points: '" & Err.Description & "'")
    
    On Error Resume Next
    
    Screen.MousePointer = vbNormal
    
    If Not mbExcelWasOpen Then
        oExcelApp.DisplayAlerts = False
        oExcelApp.Quit
    Else
        oExcelApp.Visible = True
    End If
    
    Set oExcelApp = Nothing
    Set oExcelDoc = Nothing
    Set oExcelNamesSheet = Nothing
    Set oExcelWkSht = Nothing
    
    TMS_UpdateSQDescriptionsFromXLS = False
    
    

End Function


Private Function OpenExcelForTMSUpdate() As Boolean

On Error Resume Next

    Set oExcelApp = GetObject(, "Excel.Application")
    mbExcelWasOpen = True
    If Err.number <> 0 Then
      mbExcelWasOpen = False
      Err.Clear
      Set oExcelApp = CreateObject("Excel.Application")
    End If
    
    If Err.number <> 0 Then
        Set oExcelApp = Nothing
        OpenExcelForTMSUpdate = False
        Exit Function
    Else
        OpenExcelForTMSUpdate = True
    End If
    
End Function

Private Function CellEmptyTMS(oRange As Object) As Boolean
    Dim bEmpty As Boolean
    If oRange Is Nothing Then
        bEmpty = True
    End If

    If Not bEmpty Then
        If Trim(CStr(oRange)) = "" Then
            bEmpty = True
        End If
    End If

    CellEmptyTMS = bEmpty
End Function

