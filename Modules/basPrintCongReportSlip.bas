Attribute VB_Name = "basPrintCongReportSlip"
Option Explicit
Dim CongRepCongNameXPos As Single
Dim CongRepCongNameYPos As Single
Dim CongRepMonthYearXPos As Single
Dim CongRepMonthYearYPos As Single
Dim CongRepCongNoXPos As Single
Dim CongRepCongNoYPos As Single
Dim CongRepPubFiguresYPos As Single
Dim CongRepRegFiguresYPos As Single
Dim CongRepAuxFiguresYPos As Single
Dim CongRepNoReportingXPos As Single
Dim CongRepBooksXPos As Single
Dim CongRepBrochuresXPos As Single
Dim CongRepHoursXPos As Single
Dim CongRepMagsXPos As Single
Dim CongRepRVsXPos As Single
Dim CongRepStudiesXPos As Single

Dim CongRepTopMargin As Single
Dim CongRepLeftMargin As Single
Dim CongRepFontSize As Single
Dim CongRepFontName As String
Dim CongRepPaperHeight As Single
Dim CongRepPaperWidth As Single
Dim CongRepTweakX As Single
Dim CongRepTweakY As Single

Dim SaveX As Single
Dim SaveY As Single

Dim DefaultPageHeight As Single
Dim DefaultPageWidth As Single

Dim msCongName As String
Dim msMonthYear As String
Dim msCongNo As String

Dim msPubsCount As String
Dim msPubsBooks As String
Dim msPubsBrochures As String
Dim msPubsHours As String
Dim msPubsMags As String
Dim msPubsRVs As String
Dim msPubsStudies As String

Dim msAuxCount As String
Dim msAuxBooks As String
Dim msAuxBrochures As String
Dim msAuxHours As String
Dim msAuxMags As String
Dim msAuxRVs As String
Dim msAuxStudies As String

Dim msRegCount As String
Dim msRegBooks As String
Dim msRegBrochures As String
Dim msRegHours As String
Dim msRegMags As String
Dim msRegRVs As String
Dim msRegStudies As String



Public Function PrintCongReportSlip() As Boolean

On Error GoTo ErrorTrap
Dim rstMinStats As Recordset, TheString As String, PrintSQL As String, SaveOrientation
Dim str As String, SavePapersize

    SetUpPrintParameters 'scalemode set here
    
    'get all the print fields....
    msCongName = GlobalParms.GetValue("DefaultCong", "AlphaVal", "")
    msCongNo = CStr(GlobalDefaultCong)
    
    With frmCongStats
    str = .cmbMonthCong.text
    msMonthYear = str & " " & _
                  CStr(ConvertServiceYearToNormalYear(CDate("01/" & _
                                                            str & _
                                                            "/" & _
                                                            .cmbYearCong.text)))
    msPubsCount = .txtNoPubs
    msPubsBooks = .txtPubBooks
    msPubsBrochures = .txtPubBro
    msPubsHours = .txtPubHrs
    msPubsMags = .txtPubMags
    msPubsRVs = .txtPubRVs
    msPubsStudies = .txtPubStu
    
    msRegCount = .txtNoReg
    msRegBooks = .txtRegBooks
    msRegBrochures = .txtRegBro
    msRegHours = .txtRegHrs
    msRegMags = .txtRegMags
    msRegRVs = .txtRegRVs
    msRegStudies = .txtRegStu
    
    msAuxCount = .txtNoAux
    msAuxBooks = .txtAuxBooks
    msAuxBrochures = .txtAuxBro
    msAuxHours = .txtAuxHrs
    msAuxMags = .txtAuxMags
    msAuxRVs = .txtAuxRVs
    msAuxStudies = .txtAuxStu
        
    End With
    
    DefaultPageWidth = Printer.Width
    DefaultPageHeight = Printer.Height
    
    'since we're printing in landscape, we must
    ' set Printer height to paper's width, and
    ' Printer width to papers's height
    Printer.Width = 566.929 * CongRepPaperHeight
    Printer.Height = 566.929 * CongRepPaperWidth
    
    SaveOrientation = Printer.Orientation
    
    Printer.Orientation = vbPRORLandscape
    
    PrintCongName
    PrintCongMonthYear
    PrintCongNo
    PrintCongPubCount
    PrintCongPubBooks
    PrintCongPubBrochures
    PrintCongPubHours
    PrintCongPubMags
    PrintCongPubRVs
    PrintCongPubStudies
    PrintCongAuxCount
    PrintCongAuxBooks
    PrintCongAuxBrochures
    PrintCongAuxHours
    PrintCongAuxMags
    PrintCongAuxRVs
    PrintCongAuxStudies
    PrintCongRegCount
    PrintCongRegBooks
    PrintCongRegBrochures
    PrintCongRegHours
    PrintCongRegMags
    PrintCongRegRVs
    PrintCongRegStudies
    
    Printer.EndDoc
    
    Printer.Orientation = SaveOrientation
    Printer.Width = DefaultPageWidth
    Printer.Height = DefaultPageHeight
    
    ShowMessage "The Congregation Report details have been sent to the default printer", 2000, frmCongStats
        
    PrintCongReportSlip = True
    
    Exit Function
ErrorTrap:
    MsgBox "There was a printing error. Please check printer and try again.", vbOKOnly + vbCritical, AppName
    PrintCongReportSlip = False
    Exit Function
    
End Function

Private Sub SetUpPrintParameters()

On Error GoTo ErrorTrap

    Printer.ScaleMode = vbCentimeters
     
    CongRepCongNameXPos = GlobalParms.GetValue("CongRepCongNameXPos", "NumFloat")
    CongRepCongNameYPos = GlobalParms.GetValue("CongRepCongNameYPos", "NumFloat")
    
    CongRepMonthYearXPos = GlobalParms.GetValue("CongRepMonthYearXPos", "NumFloat")
    CongRepMonthYearYPos = GlobalParms.GetValue("CongRepMonthYearYPos", "NumFloat")
    
    CongRepCongNoXPos = GlobalParms.GetValue("CongRepCongNoXPos", "NumFloat")
    CongRepCongNoYPos = GlobalParms.GetValue("CongRepCongNoYPos", "NumFloat")
    
    CongRepPubFiguresYPos = GlobalParms.GetValue("CongRepPubFiguresYPos", "NumFloat")
    CongRepRegFiguresYPos = GlobalParms.GetValue("CongRepRegFiguresYPos", "NumFloat")
    CongRepAuxFiguresYPos = GlobalParms.GetValue("CongRepAuxFiguresYPos", "NumFloat")
    CongRepNoReportingXPos = GlobalParms.GetValue("CongRepNoReportingXPos", "NumFloat")
    CongRepBooksXPos = GlobalParms.GetValue("CongRepBooksXPos", "NumFloat")
    CongRepBrochuresXPos = GlobalParms.GetValue("CongRepBrochuresXPos", "NumFloat")
    CongRepHoursXPos = GlobalParms.GetValue("CongRepHoursXPos", "NumFloat")
    CongRepMagsXPos = GlobalParms.GetValue("CongRepMagsXPos", "NumFloat")
    CongRepRVsXPos = GlobalParms.GetValue("CongRepRVsXPos", "NumFloat")
    CongRepStudiesXPos = GlobalParms.GetValue("CongRepStudiesXPos", "NumFloat")
    
    CongRepTopMargin = GlobalParms.GetValue("CongRepTopMargin", "NumFloat")
    CongRepLeftMargin = GlobalParms.GetValue("CongRepLeftMargin", "NumFloat")
    CongRepFontSize = GlobalParms.GetValue("CongRepFontSize", "NumFloat")
    CongRepFontName = GlobalParms.GetValue("CongRepFontName", "AlphaVal")
    CongRepPaperHeight = GlobalParms.GetValue("CongRepPaperHeight", "NumFloat")
    CongRepPaperWidth = GlobalParms.GetValue("CongRepPaperWidth", "NumFloat")
    CongRepTweakX = GlobalParms.GetValue("CongRepTweakX", "NumFloat")
    CongRepTweakY = GlobalParms.GetValue("CongRepTweakY", "NumFloat")
    Printer.Font.Name = CongRepFontName
    Printer.Font.Size = CongRepFontSize

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintCongName()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepCongNameXPos
    .CurrentY = CongRepCongNameYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msCongName
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongNo()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepCongNoXPos
    .CurrentY = CongRepCongNoYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msCongNo
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintCongMonthYear()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepMonthYearXPos
    .CurrentY = CongRepMonthYearYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msMonthYear
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintCongPubCount()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepNoReportingXPos
    .CurrentY = CongRepPubFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msPubsCount
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongPubBooks()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepBooksXPos
    .CurrentY = CongRepPubFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msPubsBooks
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongPubBrochures()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepBrochuresXPos
    .CurrentY = CongRepPubFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msPubsBrochures
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongPubHours()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepHoursXPos
    .CurrentY = CongRepPubFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msPubsHours
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongPubMags()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepMagsXPos
    .CurrentY = CongRepPubFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msPubsMags
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongPubRVs()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepRVsXPos
    .CurrentY = CongRepPubFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msPubsRVs
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongPubStudies()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepStudiesXPos
    .CurrentY = CongRepPubFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msPubsStudies
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
'********************************************************
Private Sub PrintCongAuxCount()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepNoReportingXPos
    .CurrentY = CongRepAuxFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msAuxCount
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongAuxBooks()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepBooksXPos
    .CurrentY = CongRepAuxFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msAuxBooks
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongAuxBrochures()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepBrochuresXPos
    .CurrentY = CongRepAuxFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msAuxBrochures
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongAuxHours()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepHoursXPos
    .CurrentY = CongRepAuxFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msAuxHours
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongAuxMags()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepMagsXPos
    .CurrentY = CongRepAuxFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msAuxMags
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongAuxRVs()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepRVsXPos
    .CurrentY = CongRepAuxFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msAuxRVs
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongAuxStudies()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepStudiesXPos
    .CurrentY = CongRepAuxFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msAuxStudies
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
'********************************************************
Private Sub PrintCongRegCount()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepNoReportingXPos
    .CurrentY = CongRepRegFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msRegCount
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongRegBooks()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepBooksXPos
    .CurrentY = CongRepRegFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msRegBooks
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongRegBrochures()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepBrochuresXPos
    .CurrentY = CongRepRegFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msRegBrochures
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongRegHours()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepHoursXPos
    .CurrentY = CongRepRegFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msRegHours
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongRegMags()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepMagsXPos
    .CurrentY = CongRepRegFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msRegMags
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongRegRVs()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepRVsXPos
    .CurrentY = CongRepRegFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msRegRVs
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintCongRegStudies()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = CongRepStudiesXPos
    .CurrentY = CongRepRegFiguresYPos - .ScaleY(CongRepFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + CongRepTweakX + CongRepLeftMargin
    .CurrentY = .CurrentY + CongRepTweakY + CongRepTopMargin
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msRegStudies
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
