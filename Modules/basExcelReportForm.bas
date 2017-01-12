Attribute VB_Name = "basExcelReportForm"
Option Explicit

'Dim oExcelApp As Excel.Application
'Dim oExcelDoc As Excel.Workbook
'Dim oExcelResultSheet As Excel.Worksheet
'Dim oExcelNamesSheet As Excel.Worksheet
'Dim oExcelWkSht As Excel.Worksheet

Dim oExcelApp As Object
Dim oExcelDoc As Object
Dim oExcelResultSheet As Object
Dim oExcelNamesSheet As Object
Dim oExcelWkSht As Object

Dim mbExcelWasOpen As Boolean



Public Function OpenExcel() As Boolean

On Error Resume Next

    Set oExcelApp = GetObject(, "Excel.Application")
    mbExcelWasOpen = True
    If Err.number <> 0 Then
      mbExcelWasOpen = False
      Err.Clear
      Set oExcelApp = CreateObject("Excel.Application")
    End If
    
    If Err.number <> 0 Then
        ShowMessage "Could not open Excel", 1500, frmCongStats
        Set oExcelApp = Nothing
        OpenExcel = False
        Exit Function
    Else
        OpenExcel = True
    End If
    
End Function

Public Sub GenerateExcelReportEntryForm(TheMonth As Long, TheNormalYear As Long)
On Error GoTo ErrorTrap
Dim i As Long, j As Long, rs As Recordset, sSQL As String, sDate1 As String, sDate2 As String
Dim bForgetAboutExcel As Boolean, sTempDate As String, colWkSht As New Collection
Dim sDateDesc As String, str As String, bOK As Boolean, oWkSht As Object, dte As Date
Dim str1 As String, str2 As String

    If Not gFSO.FolderExists(gsDocsDirectory) Then
        ShowMessage "No valid folder for documents", 1500, frmCongStats
        Exit Sub
    End If
    
    If Not OpenExcel Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    oExcelApp.Visible = False
    
    bForgetAboutExcel = True 'init
    
    'set up the document
    
    Set oExcelDoc = oExcelApp.Workbooks.Add
    Set oExcelResultSheet = oExcelDoc.Worksheets.Add
    oExcelResultSheet.Name = "CMS Report Month " & GetMonthName(TheMonth) & " " & TheNormalYear
    
    'remove the 3 default sheets
    On Error Resume Next
    oExcelApp.DisplayAlerts = False
    oExcelDoc.Worksheets("Sheet1").Delete
    oExcelDoc.Worksheets("Sheet2").Delete
    oExcelDoc.Worksheets("Sheet3").Delete
    oExcelApp.DisplayAlerts = True
    On Error GoTo ErrorTrap
    
    'get date range - start of service year to specified date...
    sDate2 = "01/" & TheMonth & "/" & TheNormalYear
    sDate1 = "01/09/" & IIf(TheMonth >= 9, TheNormalYear, TheNormalYear - 1)
        
    sTempDate = sDate1
    
    'get missing reporters from start of service until specified date
    Do Until CDate(sTempDate) > CDate(sDate2)
            
        Set rs = GetMissingReportsServiceYearRecSet(sTempDate)
        
        With rs
        
        'if there are some missing reporters this month, create the worksheet...
        If Not .BOF Then
        
            bForgetAboutExcel = False
            
            If CDate(sTempDate) = CDate(sDate2) Then
                sDateDesc = "Report Month - " & GetMonthName(Month(sTempDate)) & " " & CStr(year(sTempDate))
            Else
                sDateDesc = "Late Report - " & GetMonthName(Month(sTempDate)) & " " & CStr(year(sTempDate))
            End If
            
            colWkSht.Add oExcelApp.Worksheets.Add, sDateDesc
            
            'put in headings on new worksheet
            colWkSht(sDateDesc).Range("B1").value = "Publisher"
            colWkSht(sDateDesc).Range("C1").value = "Books"
            colWkSht(sDateDesc).Range("D1").value = "Booklets"
            colWkSht(sDateDesc).Range("E1").value = "Hours"
            colWkSht(sDateDesc).Range("F1").value = "Magazines"
            colWkSht(sDateDesc).Range("G1").value = "Return Visits"
            colWkSht(sDateDesc).Range("H1").value = "Studies"
            colWkSht(sDateDesc).Range("I1").value = "Tracts"
            colWkSht(sDateDesc).Range("J1").value = "Remarks"
    
            'hide the PersonID column
            colWkSht(sDateDesc).Columns("A").Hidden = True
            
            'hide columns J to AI which contains other hidden details
            colWkSht(sDateDesc).Columns("K:AM").Hidden = True
            
            'put some special text into column K so we know this is a CMS document...
            colWkSht(sDateDesc).Range("K2").value = "CMS Field Ministry Reporting Entry Spreadsheet"
            colWkSht(sDateDesc).Range("K3").value = "CMS Reporting Entry Sheet"
            colWkSht(sDateDesc).Range("K4").value = sDate2 'society reporting month
            colWkSht(sDateDesc).Range("K5").value = sTempDate 'actual reporting month
            
            'freeze the top row
            colWkSht(sDateDesc).Range("A2").Select
            oExcelApp.ActiveWindow.FreezePanes = True
            
            'Name the sheet
            colWkSht(sDateDesc).Name = sDateDesc
   
            i = 1
            'now for the names...
            Do Until .EOF Or .BOF
                
                colWkSht(sDateDesc).Range("A" & i + 1).value = !PersonID
                colWkSht(sDateDesc).Range("B" & i + 1).value = CongregationMember.LastFirstNameMiddleName(!PersonID)
                
                'cols L to AC contain figures split depending on whether pub/aux/reg/spec
                Select Case True
                Case CongregationMember.IsRegPio(!PersonID, CDate(sTempDate))
                    colWkSht(sDateDesc).Range("S" & i + 1).value = "=C" & i + 1 'bk
                    colWkSht(sDateDesc).Range("T" & i + 1).value = "=D" & i + 1 'bro
                    colWkSht(sDateDesc).Range("U" & i + 1).value = "=E" & i + 1 'hrs
                    colWkSht(sDateDesc).Range("V" & i + 1).value = "=F" & i + 1 'mags
                    colWkSht(sDateDesc).Range("W" & i + 1).value = "=G" & i + 1 'RVs
                    colWkSht(sDateDesc).Range("X" & i + 1).value = "=H" & i + 1 'stu
                    colWkSht(sDateDesc).Range("Y" & i + 1).value = "=I" & i + 1 'tra
                Case CongregationMember.IsAuxPio(!PersonID, CDate(sTempDate))
                    colWkSht(sDateDesc).Range("Z" & i + 1).value = "=C" & i + 1 'bk
                    colWkSht(sDateDesc).Range("AA" & i + 1).value = "=D" & i + 1 'bro
                    colWkSht(sDateDesc).Range("AB" & i + 1).value = "=E" & i + 1 'hrs
                    colWkSht(sDateDesc).Range("AC" & i + 1).value = "=F" & i + 1 'mags
                    colWkSht(sDateDesc).Range("AD" & i + 1).value = "=G" & i + 1 'RVs
                    colWkSht(sDateDesc).Range("AE" & i + 1).value = "=H" & i + 1 'stu
                    colWkSht(sDateDesc).Range("AF" & i + 1).value = "=H" & i + 1 'tra
                Case CongregationMember.IsSpecPio(!PersonID, CDate(sTempDate))
                    colWkSht(sDateDesc).Range("AG" & i + 1).value = "=C" & i + 1 'bk
                    colWkSht(sDateDesc).Range("AH" & i + 1).value = "=D" & i + 1 'bro
                    colWkSht(sDateDesc).Range("AI" & i + 1).value = "=E" & i + 1 'hrs
                    colWkSht(sDateDesc).Range("AJ" & i + 1).value = "=F" & i + 1 'mags
                    colWkSht(sDateDesc).Range("AK" & i + 1).value = "=G" & i + 1 'RVs
                    colWkSht(sDateDesc).Range("AL" & i + 1).value = "=H" & i + 1 'stu
                    colWkSht(sDateDesc).Range("AM" & i + 1).value = "=H" & i + 1 'tra
                Case Else 'must be just pub
                    colWkSht(sDateDesc).Range("L" & i + 1).value = "=C" & i + 1 'bk
                    colWkSht(sDateDesc).Range("M" & i + 1).value = "=D" & i + 1 'bro
                    colWkSht(sDateDesc).Range("N" & i + 1).value = "=E" & i + 1 'hrs
                    colWkSht(sDateDesc).Range("O" & i + 1).value = "=F" & i + 1 'mags
                    colWkSht(sDateDesc).Range("P" & i + 1).value = "=G" & i + 1 'RVs
                    colWkSht(sDateDesc).Range("Q" & i + 1).value = "=H" & i + 1 'stu
                    colWkSht(sDateDesc).Range("R" & i + 1).value = "=H" & i + 1 'tra
                End Select
                
                .MoveNext
                i = i + 1
                
            Loop
            
            'put number of rows into K6 for future use
            colWkSht(sDateDesc).Range("K6").value = i
          
            'bold headings
            colWkSht(sDateDesc).Range("B1:J1").Font.Bold = True
            
            'auto-size column widths
            colWkSht(sDateDesc).Columns("B:J").AutoFit
            
            'increase width of Remarks columns
            colWkSht(sDateDesc).Columns("J").ColumnWidth = 40
'            oExcelDoc.ActiveSheet.Columns("J").ColumnWidth = 1000
            
            'put border around
            colWkSht(sDateDesc).Range("B1", "J" & i).BorderAround , -4138 'xlMedium
              
            'unlock data-entry cells (default is 'locked'), then protect the whole sheet
            colWkSht(sDateDesc).Range("C2:J" & i).Locked = False
            colWkSht(sDateDesc).Protect
            
        End If
        
        End With
            
        sTempDate = DateAdd("m", 1, sTempDate)
    Loop
                                    
    'now add the result sheet
    
    If Not bForgetAboutExcel Then

        'Title for the page
        oExcelResultSheet.Range("A1").value = GetMonthName(Month(sDate2)) & " " & CStr(year(sDate2)) & " report for the Branch"
        
        'the result cell headings
        oExcelResultSheet.Range("B4").value = "Pubs"
        oExcelResultSheet.Range("B5").value = "Aux Pios"
        oExcelResultSheet.Range("B6").value = "Reg Pios"
        oExcelResultSheet.Range("C3").value = "No"
        oExcelResultSheet.Range("D3").value = "Books"
        oExcelResultSheet.Range("E3").value = "Booklets"
        oExcelResultSheet.Range("F3").value = "Hours"
        oExcelResultSheet.Range("G3").value = "Mags"
        oExcelResultSheet.Range("H3").value = "RVs"
        oExcelResultSheet.Range("I3").value = "Studies"
        oExcelResultSheet.Range("J3").value = "Tracts"
        oExcelResultSheet.Range("B9").value = "Active Publishers:"

        'put border around result cells
        oExcelResultSheet.Range("B3", "J6").BorderAround , -4138 'xlMedium
        oExcelResultSheet.Range("B9", "D9").BorderAround , -4138 'xlMedium
        
        'bold headings
        oExcelResultSheet.Range("B4:B6").Font.Bold = True
        oExcelResultSheet.Range("C3:J3").Font.Bold = True
        oExcelResultSheet.Range("B9").Font.Bold = True
        oExcelResultSheet.Range("A1").Font.Bold = True
        oExcelResultSheet.Range("A1").Font.Underline = True

        'put hidden text
        oExcelResultSheet.Columns("K").Hidden = True
        oExcelResultSheet.Range("K2").value = "CMS Field Ministry Reporting Entry Spreadsheet"
        oExcelResultSheet.Range("K3").value = "CMS Report Totals Sheet"
        oExcelResultSheet.Range("K4").value = sDate2 'society reporting month
        
        '****** put in the calculations...
        
        'PUB COUNTS
        'get count for no pubs for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "COUNTIF('" & oWkSht.Name & "'!N2:N" & CellVal(oWkSht.Name, "K6") & ","">0"")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("C4").value = "=SUM(" & str2 & ")"
        
        'get count for no Aux for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "COUNTIF('" & oWkSht.Name & "'!AB2:AB" & CellVal(oWkSht.Name, "K6") & ","">0"")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("C5").value = "=SUM(" & str2 & ")"
        
        'get count for no Reg for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "COUNTIF('" & oWkSht.Name & "'!U2:U" & CellVal(oWkSht.Name, "K6") & ","">0"")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("C6").value = "=SUM(" & str2 & ")"
        
        'SUM BOOKS
        'get books for pubs for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!L2:L" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("D4").value = "=SUM(" & str2 & ")"
        
        'get books for Regs for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!S2:S" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("D6").value = "=SUM(" & str2 & ")"
        
        'get books for Aux for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!Z2:Z" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("D5").value = "=SUM(" & str2 & ")"
       
        'SUM BROCHURES
        'get brochures for pubs for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!M2:M" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("E4").value = "=SUM(" & str2 & ")"
        
        'get brochures for Reg for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!T2:T" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("E6").value = "=SUM(" & str2 & ")"
        
        'get brochures for Aux for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!AA2:AA" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("E5").value = "=SUM(" & str2 & ")"
        
        'SUM HOURS
        'get hours for pubs for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!N2:N" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("F4").value = "=SUM(" & str2 & ")"
        
        'get hours for Regs for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!U2:U" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("F6").value = "=SUM(" & str2 & ")"
       
        'get hours for Aux for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!AB2:AB" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("F5").value = "=SUM(" & str2 & ")"
       
        'SUM MAGS
        'get mags for pubs for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!O2:O" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("G4").value = "=SUM(" & str2 & ")"
        
        'get mags for reg for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!V2:V" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("G6").value = "=SUM(" & str2 & ")"
       
        'get mags for aux for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!AC2:AC" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("G5").value = "=SUM(" & str2 & ")"
       
        'SUM RVs
        'get RVs for pubs for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!P2:P" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("H4").value = "=SUM(" & str2 & ")"
        
        'get RVs for Reg for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!W2:W" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("H6").value = "=SUM(" & str2 & ")"
       
        'get RVs for Aux for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!AD2:AD" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("H5").value = "=SUM(" & str2 & ")"
       
        'SUM STUDIES
        'get Studies for pubs for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!Q2:Q" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("I4").value = "=SUM(" & str2 & ")"
        
        'get Studies for reg for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!X2:X" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("I6").value = "=SUM(" & str2 & ")"
       
        'get Studies for Aux for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!AE2:AE" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("I5").value = "=SUM(" & str2 & ")"
        
               
        'SUM TRACTS
        'get Tracts for pubs for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!R2:R" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("J4").value = "=SUM(" & str2 & ")"
        
        'get Tracts for reg for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!Y2:Y" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("J6").value = "=SUM(" & str2 & ")"
       
        'get Tracts for Aux for each sheet
        j = 1
        str1 = ""
        str2 = ""
        For Each oWkSht In colWkSht
            str1 = "SUM('" & oWkSht.Name & "'!AF2:AF" & CellVal(oWkSht.Name, "K6") & ")"
            str2 = str2 & str1
            If j < colWkSht.Count Then
                str2 = str2 & ", "
            End If
            j = j + 1
        Next
        oExcelResultSheet.Range("J5").value = "=SUM(" & str2 & ")"
               
               
               
        'DONE CALCULATIONS
        
        'put in number of active pubs
        dte = DateAdd("m", -1, Now)
        dte = CDate("01/" & Month(dte) & "/" & year(dte))
        oExcelResultSheet.Range("D9").value = GetNumberActivePubsInPeriod(dte, dte)
        
        'unlock the num of active pubs cell
        oExcelResultSheet.Range("D9").Locked = False
        
        'protect the whole result sheet
        oExcelResultSheet.Protect
        
        'save the doc - first get unique file name
        bOK = False
        Do Until bOK
            str = gsDocsDirectory & "\" & GetMonthName(TheMonth) & " " & TheNormalYear & _
                    " Field Service Reports " & _
                            Replace(Replace(Now, ":", "-"), "/", "-")
            bOK = Not gFSO.FileExists(str)
        Loop
        
        oExcelDoc.SaveAs str

        Screen.MousePointer = vbNormal
        
        ShowMessage "Excel spreadsheet generated", 1200, frmCongStats
                
    Else
        ShowMessage "No publishers for this period!", 1000, frmCongStats
    End If
    
    
    'clean up and leave
    
    On Error Resume Next
    
    Screen.MousePointer = vbNormal
    
    rs.Close
    Set rs = Nothing

    For i = 1 To colWkSht.Count
'        Set colWkSht.Item(i) = Nothing
        colWkSht.Remove i
    Next i
    
    If Not bForgetAboutExcel Then
        oExcelApp.Visible = True
    Else
        If Not mbExcelWasOpen Then
            'if there's nothing to show and Excel wasn't open before, shut it now.
            oExcelApp.DisplayAlerts = False
            oExcelApp.Quit
        End If
    End If
    
    Set oExcelApp = Nothing
    Set oExcelDoc = Nothing
    Set oExcelResultSheet = Nothing
    
    Screen.MousePointer = vbNormal
    
'    Application.Windows(oExcelDoc).Activate

    Exit Sub
    
ErrorTrap:

    On Error Resume Next
     
    str = Err.Description
     
    Screen.MousePointer = vbNormal
    
    If Not mbExcelWasOpen Then
        oExcelApp.DisplayAlerts = False
        oExcelApp.Quit
    Else
        oExcelApp.Visible = True
    End If
    
    Set oExcelApp = Nothing
    Set oExcelDoc = Nothing
    Set oExcelResultSheet = Nothing

'    Call EndProgram

    MsgBox "A problem occurred while processing the spreadsheet: " & str, vbOKOnly + vbExclamation, AppName

End Sub

Public Sub GenerateExcelBulkPersonEntryForm()
On Error GoTo ErrorTrap
Dim bForgetAboutExcel As Boolean
Dim str As String, bOK As Boolean, oWkSht As Object
Dim str1 As String, str2 As String

    If Not gFSO.FolderExists(gsDocsDirectory) Then
        ShowMessage "No valid folder for documents", 1500, frmPersonalDetails
        Exit Sub
    End If
    
    If Not OpenExcel Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    oExcelApp.Visible = False
    
    'set up the document
    
    Set oExcelDoc = oExcelApp.Workbooks.Add
    Set oExcelNamesSheet = oExcelDoc.Worksheets.Add
    oExcelNamesSheet.Name = "CMS Personal Details"
    
    'remove the 3 default sheets
    On Error Resume Next
    oExcelApp.DisplayAlerts = False
    oExcelDoc.Worksheets("Sheet1").Delete
    oExcelDoc.Worksheets("Sheet2").Delete
    oExcelDoc.Worksheets("Sheet3").Delete
    oExcelApp.DisplayAlerts = True
    On Error GoTo ErrorTrap
            
    'put in headings on new worksheet
    oExcelNamesSheet.Range("A1").value = "First Name"
    oExcelNamesSheet.Range("B1").value = "Official First Name"
    oExcelNamesSheet.Range("C1").value = "Middle Name"
    oExcelNamesSheet.Range("D1").value = "Last Name"
    oExcelNamesSheet.Range("E1").value = "Gender (M/F)"
    oExcelNamesSheet.Range("F1").value = "DoB"
    oExcelNamesSheet.Range("G1").value = "Address 1"
    oExcelNamesSheet.Range("H1").value = "Address 2"
    oExcelNamesSheet.Range("I1").value = "Address 3"
    oExcelNamesSheet.Range("J1").value = "Address 4"
    oExcelNamesSheet.Range("K1").value = "Postcode"
    oExcelNamesSheet.Range("L1").value = "Home Phone"
    oExcelNamesSheet.Range("M1").value = "Mobile Phone"
    oExcelNamesSheet.Range("N1").value = "Mobile Phone 2"
    oExcelNamesSheet.Range("O1").value = "Email"
    
    oExcelNamesSheet.Range("Q500").value = "CMS"
    oExcelNamesSheet.Columns("Q").Hidden = True
            
    'freeze the top row
    oExcelNamesSheet.Range("A2").Select
    oExcelApp.ActiveWindow.FreezePanes = True
            
    'bold headings
    oExcelNamesSheet.Range("A1:O1").Font.Bold = True
        
    'auto-size column widths
    oExcelNamesSheet.Columns("A:P").AutoFit
    
    'widen some columns
    oExcelNamesSheet.Columns("A").ColumnWidth = 22
    oExcelNamesSheet.Columns("B").ColumnWidth = 22
    oExcelNamesSheet.Columns("C").ColumnWidth = 22
    oExcelNamesSheet.Columns("D").ColumnWidth = 22
    oExcelNamesSheet.Columns("E").ColumnWidth = 12
    oExcelNamesSheet.Columns("F").ColumnWidth = 10
    oExcelNamesSheet.Columns("G").ColumnWidth = 30
    oExcelNamesSheet.Columns("H").ColumnWidth = 30
    oExcelNamesSheet.Columns("I").ColumnWidth = 30
    oExcelNamesSheet.Columns("J").ColumnWidth = 30
    oExcelNamesSheet.Columns("L").ColumnWidth = 15
    oExcelNamesSheet.Columns("M").ColumnWidth = 15
    oExcelNamesSheet.Columns("N").ColumnWidth = 15
    oExcelNamesSheet.Columns("O").ColumnWidth = 30
            
    'unlock data-entry cells (default is 'locked'), then protect the whole sheet
    oExcelNamesSheet.Range("A2:P500").Locked = False
    oExcelNamesSheet.Range("A2:P500").NumberFormat = "@" 'set cell format to text
    oExcelNamesSheet.Protect
        
    'save the doc - first get unique file name
    bOK = False
    Do Until bOK
        str = gsDocsDirectory & "\CMS Personal Details Bulk Load " & _
                        Replace(Replace(Now, ":", "-"), "/", "-")
        bOK = Not gFSO.FileExists(str)
    Loop
    
    oExcelDoc.SaveAs str

    Screen.MousePointer = vbNormal
    
    ShowMessage "Excel spreadsheet generated", 1200, frmPersonalDetails
                
    
    'clean up and leave
    
    On Error Resume Next
    
    Screen.MousePointer = vbNormal
    
    oExcelApp.Visible = True
    
    Set oExcelApp = Nothing
    Set oExcelDoc = Nothing
    Set oExcelNamesSheet = Nothing
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
    
ErrorTrap:

    On Error Resume Next
     
    str = Err.Description
     
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

    MsgBox "A problem occurred while processing the spreadsheet: " & str, vbOKOnly + vbExclamation, AppName

End Sub

Private Function GetMissingReportsServiceYearRecSet(NormalDate As String, _
                                                    Optional InclLongTermInactive As Boolean = True, _
                                                    Optional InclSpecPios As Boolean = True) As Recordset
On Error GoTo ErrorTrap

Dim TheString As String, DateString_US As String, Date2 As String
Dim lsNormalYear As String
            
    
    'Construct SQL to give us names of publishers that have not reported
    ' this month....
    TheString = "SELECT a.PersonID, b.LastName & ', ' & b.FirstName & ' ' & b.MiddleName " & _
                "FROM tblPublisherDates a INNER JOIN tblNameAddress b " & _
                "       ON a.PersonID = b.ID " & _
                "WHERE StartDate <= " & GetDateStringForSQLWhere(NormalDate) & _
                " AND EndDate >= " & GetDateStringForSQLWhere(NormalDate) & _
                " AND StartReason <> 2 " & _
                "AND a.PersonID NOT IN " & _
                "            (SELECT PersonID " & _
                            " FROM tblMinReports " & _
                            " WHERE ActualMinPeriod = " & GetDateStringForSQLWhere(NormalDate) & ") "
                            
    If Not InclSpecPios Then
        TheString = TheString & " AND PersonID NOT IN (SELECT PersonID " & _
                                                     "FROM tblSpecPioDates " & _
                                                     "WHERE StartDate <= " & GetDateStringForSQLWhere(NormalDate) & _
                                                    " AND EndDate >= " & GetDateStringForSQLWhere(NormalDate) & ") "
    End If
    
    TheString = TheString & " ORDER BY 2 "
                
    Set GetMissingReportsServiceYearRecSet = CMSDB.OpenRecordset(TheString, dbOpenDynaset)

    Exit Function
ErrorTrap:
    EndProgram
    
End Function
Private Function CellVal(SheetName As String, Coord As String) As Variant
On Error GoTo ErrorTrap

    CellVal = oExcelDoc.Worksheets(SheetName).Range(Coord)

    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Sub GetAndValidateExcelReport(BranchReportPeriod As String)
On Error GoTo ErrorTrap
Dim sExcelPath As String, str As String, sDate1 As String, sDate2 As String, i As Long, sDateCur As String
Dim rs As Recordset
Dim rs2 As Recordset
Dim bSaveReport As Boolean
Dim bReportFound As Boolean

Dim lPersonID As Long
Dim sPersonID As String
Dim lMaxRow As Long
Dim sHours As String
Dim sMags As String
Dim sRVs As String
Dim sStudies As String
Dim sBooks As String
Dim sBrochures As String
Dim sTracts As String
Dim sRemarks As String
Dim lHours As Double
Dim lMags As Long
Dim lRVs As Long
Dim lStudies As Long
Dim lBooks As Long
Dim lBrochures As Long
Dim lTracts As Long
Dim sHoursOld As String
Dim sMagsOld As String
Dim sRVsOld As String
Dim sStudiesOld As String
Dim sBooksOld As String
Dim sBrochuresOld As String
Dim sRemarksOld As String
Dim sTractsOld As String
Dim lHoursOld As Double
Dim lMagsOld As Long
Dim lRVsOld As Long
Dim lStudiesOld As Long
Dim lBooksOld As Long
Dim lBrochuresOld As Long
Dim lTractsOld As Long

    If Not gFSO.FolderExists(gsDocsDirectory) Then
        ShowMessage "No valid folder for documents", 1500, frmCongStats
        GoTo GetOut
    End If
    
    If Not OpenExcel Then
        GoTo GetOut
    End If

    bReportFound = False 'init
    
    Screen.MousePointer = vbHourglass
    
    
    sDate1 = "01/09/" & IIf(Month(BranchReportPeriod) >= 9, year(BranchReportPeriod), year(BranchReportPeriod) - 1)
    
    'allow user to select the file
    If Not GetFile(sExcelPath, _
                   "C.M.S. Open Report Excel Spreadsheet", _
                    frmCongStats.CommonDialog1, _
                    gsDocsDirectory, _
                     "", "Microsoft Excel Files (2007+)|*.xlsx|Microsoft Excel Files (97, 2000, 2003)|*.xls", False, "xlsx") Then
                    
        ShowMessage "No file selected", 1500, frmCongStats, , , True
        GoTo GetOut
    End If
    
    'open the spreadsheet
    On Error Resume Next
    oExcelApp.Visible = False
    Set oExcelDoc = oExcelApp.Application.Workbooks.Open(sExcelPath)
    If Err.number <> 0 Then
        MsgBox "Could not open spreadsheet", vbOKOnly + vbInformation, AppName
        GoTo GetOut
    End If
    On Error GoTo ErrorTrap
    
    'check that all sheets relate to CMS and relate to specified reporting period and report nt already
    ' sent to branch
    For Each oExcelWkSht In oExcelDoc.Worksheets
        
        Select Case oExcelWkSht.Range("K2").value
        Case "CMS Field Ministry Reporting Entry Spreadsheet"
        Case Else
            MsgBox "Invalid worksheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
            GoTo GetOut
        End Select
    
        Select Case oExcelWkSht.Range("K3").value
        Case "CMS Reporting Entry Sheet"
            str = oExcelWkSht.Range("K4").value
            If Not IsDate(str) Then
                MsgBox "Invalid Branch reporting date on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                GoTo GetOut
            End If
            str = Format(str, "mm/dd/yyyy")
            If Not ValidDate(str) Then
                MsgBox "Invalid Branch reporting date on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                GoTo GetOut
            End If
            If MinReportSentToBranch(CDate(str)) Then
                MsgBox "Report already marked as sent to Branch for " & Format(str, "mmmm yyyy") & " - sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                GoTo GetOut
            End If
            str = oExcelWkSht.Range("K5").value
            If Not IsDate(str) Then
                MsgBox "Invalid reporting date on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                GoTo GetOut
            End If
            str = Format(str, "mm/dd/yyyy")
            If Not ValidDate(str) Then
                MsgBox "Invalid reporting date on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                GoTo GetOut
            End If
        Case "CMS Report Totals Sheet"
            str = oExcelWkSht.Range("K4").value
            If Not IsDate(str) Then
                MsgBox "Invalid Branch reporting date on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                GoTo GetOut
            End If
            str = Format(str, "mm/dd/yyyy")
            If Not ValidDate(str) Then
                MsgBox "Invalid Branch reporting date on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                GoTo GetOut
            End If
            If MinReportSentToBranch(CDate(str)) Then
                MsgBox "Report already marked as sent to Branch for " & Format(str, "mmmm yyyy") & " - sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                GoTo GetOut
            End If
        Case Else
            MsgBox "Invalid worksheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
            GoTo GetOut
        End Select
        
        Select Case oExcelWkSht.Range("K3").value
        Case "CMS Reporting Entry Sheet"
            sDate2 = Format(oExcelWkSht.Range("K4").value, "mm/dd/yyyy")
            sDateCur = Format(oExcelWkSht.Range("K5").value, "mm/dd/yyyy")
            If CDate(sDateCur) < CDate(sDate1) Or CDate(sDateCur) > CDate(sDate2) Then
                MsgBox "Report date out of range on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                GoTo GetOut
            End If
            If CDate(BranchReportPeriod) <> CDate(sDate2) Then
                MsgBox "Branch reporting date not equal to that specified - sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                GoTo GetOut
            End If
            If Not IsNumber(oExcelWkSht.Range("K6").value, False, False, False) Then
                MsgBox "Invalid MaxRow value on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                GoTo GetOut
            End If
            lMaxRow = CLng(oExcelWkSht.Range("K6").value)
            
            For i = 2 To lMaxRow 'for each publisher on sheet
                sPersonID = oExcelWkSht.Range("A" & i).value
                If Not IsNumber(sPersonID) Then
                    MsgBox "Invalid PersonID on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                    GoTo GetOut
                End If
                lPersonID = CLng(sPersonID)
                If Not CongregationMember.IsPublisher(lPersonID, CDate(sDateCur)) Then
                    MsgBox "Non-publisher on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                    GoTo GetOut
                End If
                
                sHours = oExcelWkSht.Range("E" & i).value
                sMags = oExcelWkSht.Range("F" & i).value
                sRVs = oExcelWkSht.Range("G" & i).value
                sStudies = oExcelWkSht.Range("H" & i).value
                sBooks = oExcelWkSht.Range("C" & i).value
                sBrochures = oExcelWkSht.Range("D" & i).value
                sTracts = oExcelWkSht.Range("I" & i).value
                sRemarks = oExcelWkSht.Range("J" & i).value
                
                If sHours <> "" Then
                    If Not IsNumber(sHours, True) Then
                        MsgBox "Invalid hours on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                        GoTo GetOut
                    End If
                    lHours = CDbl(sHours)
                Else
                    lHours = 0
                End If
                If sMags <> "" Then
                    If Not IsNumber(sMags) Then
                        MsgBox "Invalid Magazines on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                        GoTo GetOut
                    End If
                    lMags = CLng(sMags)
                Else
                    lMags = 0
                End If
                If sRVs <> "" Then
                    If Not IsNumber(sRVs) Then
                        MsgBox "Invalid Return Visits on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                        GoTo GetOut
                    End If
                    lRVs = CLng(sRVs)
                Else
                    lRVs = 0
                End If
                If sStudies <> "" Then
                    If Not IsNumber(sStudies) Then
                        MsgBox "Invalid Studies on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                        GoTo GetOut
                    End If
                    lStudies = CLng(sStudies)
                Else
                    lStudies = 0
                End If
                If sBooks <> "" Then
                    If Not IsNumber(sBooks) Then
                        MsgBox "Invalid Books on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                        GoTo GetOut
                    End If
                    lBooks = CLng(sBooks)
                Else
                    lBooks = 0
                End If
                If sBrochures <> "" Then
                    If Not IsNumber(sBrochures) Then
                        MsgBox "Invalid Brochures on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                        GoTo GetOut
                    End If
                    lBrochures = CLng(sBrochures)
                Else
                    lBrochures = 0
                End If
                
                If sTracts <> "" Then
                    If Not IsNumber(sTracts) Then
                        MsgBox "Invalid Tracts on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                        GoTo GetOut
                    End If
                    lTracts = CLng(sBrochures)
                Else
                    lTracts = 0
                End If
                
                If lHours = 0 And _
                    (lBooks > 0 Or lBrochures > 0 Or lStudies > 0 Or lRVs > 0 Or lMags > 0) Then
                    MsgBox "Zero hours for non-zero books/brochures/magazines/return visits/studies on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                    GoTo GetOut
                End If
                
                If lRVs < lStudies Then
                    MsgBox "Return Visits should be at least equal to number of Studies on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                    GoTo GetOut
                End If
            
                If CLng(CongregationMember.InfirmityLevel(CInt(lPersonID))) <= _
                        GlobalParms.GetValue("ThresholdForReportIn15MinInc", "NumVal") Then
                        
                    If GetFractionPart(lHours) > 0 Then
                        If MsgBox(CongregationMember.NameWithMiddleInitial(lPersonID) & _
                            " is not currently authorised to report in 15 minute increments. " & _
                            "Do you want to allow this? (Sheet '" & oExcelWkSht.Name & "')", vbYesNo + vbExclamation, AppName) = vbNo Then
                            GoTo GetOut
                        End If
                    End If
                End If
                    
                If GetFractionPart(lHours) > 0 Then
                    Select Case GetFractionPart(lHours)
                    Case 0
                    Case 0.25, 0.75
                    Case 0.5
                    Case Else
                        MsgBox "Invalid decimal hours on sheet '" & oExcelWkSht.Name & "'", vbOKOnly + vbInformation, AppName
                        GoTo GetOut
                    End Select
                End If
                
                If lHours = 0 And sHours <> "" And Len(Trim(sRemarks)) = 0 Then
                    MsgBox "If report has zero hours you should enter some remarks (Sheet '" & oExcelWkSht.Name & "')", vbOKOnly + vbExclamation, AppName
                    GoTo GetOut
                End If
            
           Next i 'next person on sheet
            
        End Select
        
    Next 'next sheet
    
    Set rs = CMSDB.OpenRecordset("tblMinReports", dbOpenDynaset)
    
    With rs
    
    'now load the reports to CMS...
    For Each oExcelWkSht In oExcelDoc.Worksheets
    
     If oExcelWkSht.Range("k3").value = "CMS Reporting Entry Sheet" Then
     
       lMaxRow = CLng(oExcelWkSht.Range("K6").value)
       
       For i = 2 To lMaxRow 'for each publisher on sheet
        
        lPersonID = CLng(oExcelWkSht.Range("A" & i).value)
        sDateCur = Format(oExcelWkSht.Range("K5").value, "mm/dd/yyyy")
        sHours = oExcelWkSht.Range("E" & i).value
        sMags = oExcelWkSht.Range("F" & i).value
        sRVs = oExcelWkSht.Range("G" & i).value
        sStudies = oExcelWkSht.Range("H" & i).value
        sBooks = oExcelWkSht.Range("C" & i).value
        sBrochures = oExcelWkSht.Range("D" & i).value
        sTracts = oExcelWkSht.Range("I" & i).value
        sRemarks = oExcelWkSht.Range("J" & i).value
    
        If sHours <> "" Then 'a report has been entered on the sheet for this person
            
            If sHours <> "" Then
                lHours = CDbl(sHours)
            Else
                lHours = 0
            End If
            If sMags <> "" Then
                lMags = CLng(sMags)
            Else
                lMags = 0
            End If
            If sRVs <> "" Then
                lRVs = CLng(sRVs)
            Else
                lRVs = 0
            End If
            If sStudies <> "" Then
                lStudies = CLng(sStudies)
            Else
                lStudies = 0
            End If
            If sBooks <> "" Then
                lBooks = CLng(sBooks)
            Else
                lBooks = 0
            End If
            If sBrochures <> "" Then
                lBrochures = CLng(sBrochures)
            Else
                lBrochures = 0
            End If
            If sTracts <> "" Then
                lTracts = CLng(sTracts)
            Else
                lTracts = 0
            End If
                        
            .FindFirst "PersonID = " & lPersonID & " AND ActualMinPeriod = " & GetDateStringForSQLWhere(sDateCur)
            If Not .NoMatch Then
                If lHours <> !NoHours Or _
                   lMags <> !NoMagazines Or _
                   lRVs <> !NoReturnVisits Or _
                   lStudies <> !NoStudies Or _
                   lBooks <> !NoBooks Or _
                   lBrochures <> !NoBooklets Or _
                   lTracts <> !NoTracts Or _
                   sRemarks <> !Remarks Then 'report already exists, new one is different
                   
                    If MsgBox("Replace " & AddApostropheToPersonName(CongregationMember.NameWithMiddleInitial(lPersonID)) & _
                              " " & Format(sDateCur, "mmmm yyyy") & " report consisting of: " & _
                              vbCrLf & vbCrLf & "Books: " & !NoBooks & _
                              vbCrLf & "Brochures: " & !NoBooklets & _
                              vbCrLf & "Hours: " & !NoHours & _
                              vbCrLf & "Magazines: " & !NoMagazines & _
                              vbCrLf & "Return Visits: " & !NoReturnVisits & _
                              vbCrLf & "Studies: " & !NoStudies & _
                              vbCrLf & "Tracts: " & !NoTracts & _
                              vbCrLf & "Remarks: '" & !Remarks & "'" & vbCrLf & vbCrLf & " with " & _
                              vbCrLf & vbCrLf & _
                              vbCrLf & "Books: " & lBooks & _
                              vbCrLf & "Brochures: " & lBrochures & _
                              vbCrLf & "Hours: " & lHours & _
                              vbCrLf & "Magazines: " & lMags & _
                              vbCrLf & "Return Visits: " & lRVs & _
                              vbCrLf & "Studies: " & lStudies & _
                              vbCrLf & "Tracts: " & lTracts & _
                              vbCrLf & "Remarks: '" & sRemarks & "'?", _
                              vbYesNo + vbQuestion + vbDefaultButton2, AppName) = vbYes Then
                              
                        bSaveReport = True
                        .Edit
                    Else
                        bSaveReport = False
                    End If
                Else
                    bSaveReport = False
                End If
            Else
                'report doesn't exist for this month and person
                bSaveReport = True
                .AddNew
                !PersonID = lPersonID
                !ActualMinPeriod = CDate(sDateCur)
                !OtherComments = ""
            End If
            
            If bSaveReport Then
            
                If CongregationMember.IsAuxPio(lPersonID, CDate(sDateCur)) Then
                    If lHours < GlobalParms.GetValue("MonthlyAuxPioHours", "NumVal") Then
                        MsgBox "Check " & Format(sDateCur, "mmmm yyyy") & " hours (" & lHours & _
                            ") for auxilliary pioneer " & _
                                CongregationMember.NameWithMiddleInitial(lPersonID), _
                                  vbOKOnly + vbInformation, AppName
                    End If
                    
                    
                    If InStr(1, sRemarks, "aux", vbTextCompare) > 0 And _
                        InStr(1, sRemarks, "pio", vbTextCompare) > 0 Then
                        
                    Else
                        If MsgBox("Do you want to add a comment indicating that " & _
                                CongregationMember.NameWithMiddleInitial(lPersonID) & _
                                " auxilliary pioneered in " & Format(sDateCur, "mmmm yyyy") & "?", _
                                  vbYesNo + vbQuestion, AppName) = vbYes Then
                            sRemarks = sRemarks & " Aux Pio."
                        End If
                    End If
                End If
                
                If CongregationMember.IsRegPio(lPersonID, CDate(sDateCur)) Then
                    If lHours < (GlobalParms.GetValue("AnnualRegPioHours", "NumVal") / 12) Then
                        MsgBox "Check " & Format(sDateCur, "mmmm yyyy") & " hours (" & lHours & _
                            ") for regular pioneer " & _
                                CongregationMember.NameWithMiddleInitial(lPersonID), _
                                  vbOKOnly + vbInformation, AppName
                    End If
                End If
                
                If CongregationMember.MovedInThisMonth(lPersonID, CDate(sDateCur)) Then
                    If InStr(1, sRemarks, "moved", vbTextCompare) > 0 And _
                        InStr(1, sRemarks, "in", vbTextCompare) > 0 Then
                        
                    Else
                        If MsgBox("Do you want to add a comment indicating that " & _
                                CongregationMember.NameWithMiddleInitial(lPersonID) & _
                                " moved in this month?", _
                                  vbYesNo + vbQuestion, AppName) = vbYes Then
                            sRemarks = sRemarks & " Moved in."
                        End If
                    End If
                End If
                
                If CongregationMember.MovedOutThisMonth(lPersonID, CDate(sDateCur)) Then
                    If InStr(1, sRemarks, "moved", vbTextCompare) > 0 And _
                        InStr(1, sRemarks, "out", vbTextCompare) > 0 Then
                        
                    Else
                        If MsgBox("Do you want to add a comment indicating that " & _
                                CongregationMember.NameWithMiddleInitial(lPersonID) & _
                                " moved out this month?", _
                                  vbYesNo + vbQuestion, AppName) = vbYes Then
                            sRemarks = sRemarks & " Moved out."
                        End If
                    End If
                End If
            
                !MinistryDoneInMonth = Month(sDateCur)
                !MinistryDoneInYear = year(ConvertNormalDateToServiceDate(CDate(sDateCur)))
                !SocietyReportingMonth = Month(sDate2)
                !SocietyReportingYear = year(ConvertNormalDateToServiceDate(CDate(sDate2)))
                !SocietyReportingPeriod = CDate(sDate2)
                !NoBooklets = lBrochures
                !NoBooks = lBooks
                !NoHours = lHours
                !NoMagazines = lMags
                !NoReturnVisits = lRVs
                !NoStudies = lStudies
                !NoTracts = lTracts
                !Remarks = sRemarks
                .Update
                bReportFound = True
                RefreshMinistryStatusForPerson lPersonID
                CongregationMember.HighlightInactivity lPersonID, CDate(sDateCur)
                
                'now insert pub rec printed flag
                Set rs2 = CMSDB.OpenRecordset("tblPubRecCardRowPrinted", dbOpenDynaset)
                With rs2
                .FindFirst "PersonID = " & lPersonID & " AND ActualMinPeriod = " & GetDateStringForSQLWhere(sDateCur)
                If .NoMatch Then
                    .AddNew
                    !ActualMinPeriod = CDate(sDateCur)
                    !PersonID = lPersonID
                    !Printed = False
                    .Update
                End If
                End With

            End If
            
        End If
        
       Next i 'next person on sheet
        
     End If
    
    Next 'next sheet (ie month)
    
    End With

    frmCongStats.UpdateCongStatsForm
    
    If bReportFound Then
        MsgBox "Report loaded successfully from Excel spreadsheet", vbOKOnly + vbInformation, AppName
    Else
        MsgBox "No valid reports found on Excel spreadsheet", vbOKOnly + vbInformation, AppName
    End If
    
GetOut:
    
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
    Set oExcelResultSheet = Nothing
    Set oExcelWkSht = Nothing
    
    rs.Close
    Set rs = Nothing
    rs2.Close
    Set rs2 = Nothing

    Exit Sub
ErrorTrap:
    On Error Resume Next
    
    str = Err.Description
     
    Screen.MousePointer = vbNormal
    
    If Not mbExcelWasOpen Then
        oExcelApp.DisplayAlerts = False
        oExcelApp.Quit
    Else
        oExcelApp.Visible = True
    End If
    
    Set oExcelApp = Nothing
    Set oExcelDoc = Nothing
    Set oExcelResultSheet = Nothing
    Set oExcelWkSht = Nothing
'    Call EndProgram
    MsgBox "A problem occurred while processing the spreadsheet: " & str, vbOKOnly + vbExclamation, AppName
End Sub


Public Sub GetAndValidateExcelPersonalDetails()
On Error GoTo ErrorTrap
Dim sExcelPath As String, str As String, i As Long, j As Long
Dim rs As Recordset
Dim rs2 As Recordset
Dim bSaveName As Boolean
Dim sFirstNameXLS As String
Dim sFirstName As String
Dim sOfficialFirstName As String
Dim sMiddleName As String
Dim sSurname As String
Dim sGenderMF As String
Dim sDOB As String
Dim sAddress1 As String
Dim sAddress2 As String
Dim sAddress3 As String
Dim sAddress4 As String
Dim sAddress5 As String
Dim sPostcode As String
Dim sHomePhone As String
Dim sMobile As String
Dim sMobile2 As String
Dim sEmail As String

    If Not gFSO.FolderExists(gsDocsDirectory) Then
        ShowMessage "No valid folder for documents", 1500, frmCongStats
        GoTo GetOut
    End If
    
    If Not OpenExcel Then
        GoTo GetOut
    End If

    Screen.MousePointer = vbHourglass
    
    'allow user to select the file
    If Not GetFile(sExcelPath, _
                   "C.M.S. Open Personal Details Excel Spreadsheet", _
                    frmPersonalDetails.CommonDialog1, _
                    gsDocsDirectory, _
                     "", "Microsoft Excel Files|*.xls", False, "xls") Then
                    
        MsgBox "Could not locate/open spreadsheet", vbOKOnly + vbInformation, AppName
        GoTo GetOut
    End If
    
    'open the spreadsheet
    On Error Resume Next
    oExcelApp.Visible = False
    Set oExcelDoc = oExcelApp.Application.Workbooks.Open(sExcelPath)
    If Err.number <> 0 Then
        MsgBox "Could not open spreadsheet", vbOKOnly + vbInformation, AppName
        GoTo GetOut
    End If
    On Error GoTo ErrorTrap
    
    'check that doc relates to CMS
    If oExcelDoc.Worksheets.Count <> 1 Then
        ShowMessage "Invalid spreadsheet", 1500, frmPersonalDetails
        GoTo GetOut
    End If
    
    Set oExcelNamesSheet = oExcelDoc.Worksheets(1)
    
    If oExcelNamesSheet.Name <> "CMS Personal Details" Then
        ShowMessage "Invalid spreadsheet", 1500, frmPersonalDetails
        GoTo GetOut
    End If
    
    If oExcelNamesSheet.Range("Q500").value <> "CMS" Then
        ShowMessage "Invalid spreadsheet", 1500, frmPersonalDetails
        GoTo GetOut
    End If
    
    sFirstNameXLS = oExcelNamesSheet.Range("A2").value
    i = 2
    
    'check each entered name on sheet
     Do While sFirstNameXLS <> "" 'for each publisher on sheet
         
         
         If oExcelNamesSheet.Range("A" & i).value = "" Then 'firstname
             MsgBox "No first name provided on row " & i, vbOKOnly + vbInformation, AppName
             GoTo GetOut
         End If
         If oExcelNamesSheet.Range("D" & i).value = "" Then 'lastname
             MsgBox "No surname provided on row " & i, vbOKOnly + vbInformation, AppName
             GoTo GetOut
         End If
         If Len(oExcelNamesSheet.Range("A" & i).value) > 100 Then 'firstname
             MsgBox "First name longer than 100 characters on row " & i, vbOKOnly + vbInformation, AppName
             GoTo GetOut
         End If
         If Len(oExcelNamesSheet.Range("B" & i).value) > 100 Then 'official firstname
             MsgBox "Official First name longer than 100 characters on row " & i, vbOKOnly + vbInformation, AppName
             GoTo GetOut
         End If
         If Len(oExcelNamesSheet.Range("C" & i).value) > 100 Then 'middle name
             MsgBox "Middle name longer than 100 characters on row " & i, vbOKOnly + vbInformation, AppName
             GoTo GetOut
         End If
         If Len(oExcelNamesSheet.Range("D" & i).value) > 100 Then 'surname
             MsgBox "Surname longer than 100 characters on row " & i, vbOKOnly + vbInformation, AppName
             GoTo GetOut
         End If
         
         If oExcelNamesSheet.Range("E" & i).value <> "M" And _
            oExcelNamesSheet.Range("E" & i).value <> "F" Then 'genderMF
             MsgBox "Gender should be M or F on row " & i, vbOKOnly + vbInformation, AppName
             GoTo GetOut
         End If
         
        If oExcelNamesSheet.Range("F" & i).value <> "" Then
            If Len(oExcelNamesSheet.Range("F" & i).value) <> 10 Then
                MsgBox "Date of birth should be formatted dd/mm/yyyy on row " & i, vbOKOnly + vbInformation, AppName
                GoTo GetOut
            End If
            
            If Not ValidDate(oExcelNamesSheet.Range("F" & i).value) Then
                MsgBox "Invalid Date of birth on row " & i, vbOKOnly + vbInformation, AppName
                GoTo GetOut
            End If
        End If
        
        If Len(oExcelNamesSheet.Range("G" & i).value) > 100 Or _
            Len(oExcelNamesSheet.Range("H" & i).value) > 100 Or _
            Len(oExcelNamesSheet.Range("I" & i).value) > 100 Or _
            Len(oExcelNamesSheet.Range("J" & i).value) > 100 Or _
            Len(oExcelNamesSheet.Range("K" & i).value) > 20 Then
             MsgBox "Address fields should be no longer than 100 characters, postcode 20 characters on row " & i, vbOKOnly + vbInformation, AppName
             GoTo GetOut
        End If
         
        If Len(oExcelNamesSheet.Range("L" & i).value) > 50 Or _
            Len(oExcelNamesSheet.Range("M" & i).value) > 50 Or _
            Len(oExcelNamesSheet.Range("N" & i).value) > 50 Then
             MsgBox "Phone fields should be no longer than 50 characters on row " & i, vbOKOnly + vbInformation, AppName
             GoTo GetOut
        End If
               
        If Len(oExcelNamesSheet.Range("O" & i).value) > 100 Then
             MsgBox "Email field should be no longer than 100 characters on row " & i, vbOKOnly + vbInformation, AppName
             GoTo GetOut
        End If
        
        If Not ValidateEmailAddress(oExcelNamesSheet.Range("O" & i).value) Then
             MsgBox "Invalid email address on row " & i, vbOKOnly + vbInformation, AppName
             GoTo GetOut
        End If
                     
        i = i + 1
        sFirstNameXLS = oExcelNamesSheet.Range("A" & i).value
     
    Loop 'next person on sheet
    
    If i = 2 Then
        ShowMessage "Nothing to load", 1500, frmPersonalDetails
        GoTo GetOut
    End If

    'now update tblNameAddress
    For j = 2 To i - 1
        
        sFirstName = oExcelNamesSheet.Range("A" & j).value
        sOfficialFirstName = oExcelNamesSheet.Range("B" & j).value
        If sOfficialFirstName = "" Then
            sOfficialFirstName = sFirstName
        End If
        sMiddleName = oExcelNamesSheet.Range("C" & j).value
        sSurname = oExcelNamesSheet.Range("D" & j).value
        sGenderMF = oExcelNamesSheet.Range("E" & j).value
        sDOB = oExcelNamesSheet.Range("F" & j).value
        sAddress1 = oExcelNamesSheet.Range("G" & j).value
        sAddress2 = oExcelNamesSheet.Range("H" & j).value
        sAddress3 = oExcelNamesSheet.Range("I" & j).value
        sAddress4 = oExcelNamesSheet.Range("J" & j).value
        sPostcode = oExcelNamesSheet.Range("K" & j).value
        sHomePhone = oExcelNamesSheet.Range("L" & j).value
        sMobile = oExcelNamesSheet.Range("M" & j).value
        sMobile2 = oExcelNamesSheet.Range("N" & j).value
        sEmail = oExcelNamesSheet.Range("O" & j).value
       
        CongregationMember.AddPersonToCMS sFirstName, sMiddleName, sSurname, sGenderMF, sDOB, 0, _
                                          sAddress1, sAddress2, sAddress3, sAddress4, sPostcode, _
                                          sHomePhone, "", sMobile, sMobile2, sEmail, True, 0, False, sOfficialFirstName, 0

        
    Next j 'next person on sheet
           
    CloseAllOpenForms False, False
    MsgBox "Personal Details loaded successfully from Excel spreadsheet", vbOKOnly + vbInformation, AppName
    frmPersonalDetails.Show vbModeless, frmMainMenu
    
GetOut:
    
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
    
    rs.Close
    Set rs = Nothing
    rs2.Close
    Set rs2 = Nothing

    Exit Sub
ErrorTrap:
    On Error Resume Next
    
    str = Err.Description
     
    Screen.MousePointer = vbNormal
    
    If Not mbExcelWasOpen Then
        oExcelApp.DisplayAlerts = False
        oExcelApp.Quit
    Else
        oExcelApp.Visible = True
    End If
    
    Set oExcelApp = Nothing
    Set oExcelDoc = Nothing
    Set oExcelResultSheet = Nothing
    Set oExcelWkSht = Nothing
'    Call EndProgram
    MsgBox "A problem occurred while processing the spreadsheet: " & str, vbOKOnly + vbExclamation, AppName
End Sub



