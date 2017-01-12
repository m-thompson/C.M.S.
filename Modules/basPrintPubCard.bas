Attribute VB_Name = "basPrintPubCard"
Option Explicit

Dim PubCardPublisherNameXPos As Single
Dim PubCardPublisherNameYPos As Single
Dim PubCardAddressXPos As Single
Dim PubCardAddressYPos As Single
Dim PubCardPublisherNameMaxWidth As Single
Dim PubCardAddressMaxWidth As Single
Dim PubCardTelNoXPos As Single
Dim PubCardTelNoYPos As Single
Dim PubCardTelNoMaxWidth As Single
Dim PubCardElderXPos As Single
Dim PubCardElderYPos As Single
Dim PubCardServantXPos As Single
Dim PubCardServantYPos As Single
Dim PubCardRegPioXPos As Single
Dim PubCardRegPioYPos As Single
Dim PubCardBaptDateXPos As Single
Dim PubCardBaptDateYPos As Single
Dim PubCardDOBXPos As Single
Dim PubCardDOBYPos As Single
Dim PubCardBaptDateMaxWidth As Single
Dim PubCardPioNoXPos As Single
Dim PubCardPioNoYPos As Single
Dim PubCardPioNoMaxWidth As Single
Dim PubCardAnointedXPos As Single
Dim PubCardAnointedYPos As Single
Dim PubCardAnointedMaxWidth As Single

Dim PubCardServiceYearXPos As Single
Dim PubCardServiceYearYPos As Single
Dim PubCardBooksXPos As Single
Dim PubCardBrochuresXPos As Single
Dim PubCardHoursXPos As Single
Dim PubCardSubscripXPos As Single
Dim PubCardMagsXPos As Single
Dim PubCardRVsXPos As Single
Dim PubCardStudiesXPos As Single
Dim PubCardRemarksXPos As Single
Dim PubCardSeptYPos As Single
Dim PubCardOctYPos As Single
Dim PubCardNovYPos As Single
Dim PubCardDecYPos As Single
Dim PubCardJanYPos As Single
Dim PubCardFebYPos As Single
Dim PubCardMarYPos As Single
Dim PubCardAprYPos As Single
Dim PubCardMayYPos As Single
Dim PubCardJunYPos As Single
Dim PubCardJulYPos As Single
Dim PubCardAugYPos As Single
Dim PubCardTotalYPos As Single
Dim PubCardRemarksMaxWidth As Single

Dim PubCardMaleXPos As Single
Dim PubCardMaleYPos As Single
Dim PubCardFemaleXPos As Single
Dim PubCardFemaleYPos As Single
Dim PubCardMobileNoXPos As Single
Dim PubCardMobileNoYPos As Single
Dim PubCardMobileNoMaxWidth As Single


Dim PubCardTopMargin As Single
Dim PubCardLeftMargin As Single
Dim PubCardFontSize As Single
Dim PubCardFontName As String
Dim PubCardPaperHeight As Single
Dim PubCardPaperWidth As Single
Dim PubCardTweakX As Single
Dim PubCardTweakY As Single

Dim SaveX As Single
Dim SaveY As Single

Dim DefaultPageHeight As Single
Dim DefaultPageWidth As Single

Dim msName As String
Dim msAddress As String
Dim msTelNo As String
Dim msIsElder As String
Dim msIsMS As String
Dim msIsPio As String
Dim msBaptDate As String
Dim msDOB As String
Dim msPioNo As String
Dim msAnointed As String
Dim tempvar As Variant
Dim TempVar2 As Boolean
Dim msYear As String
Dim msBooks As String
Dim msBrochures As String
Dim msHours As String
Dim msMags As String
Dim msRVs As String
Dim msStudies As String
Dim msRemarks As String
Dim msMobileNo As String
Dim msIsMale As String
Dim msIsFemale As String

Public Function PrintPublisherCard(ByVal PersonID As Long, _
                                   ByVal PrintHeader As Boolean, _
                                   ByVal PrintYear As Boolean, _
                                   ByVal PrintStats As Boolean, _
                                   ByVal UnprintedOnly As Boolean, _
                                   ByVal PrintTotals As Boolean, _
                                   ByVal ServiceYear As String, _
                                   ByVal CardSide As Long) As Boolean

On Error GoTo ErrorTrap
Dim rstMinStats As Recordset, TheString As String, PrintSQL As String
Dim bNoStats As Boolean, lCardType As Long

    SetUpPrintParameters PersonID, CardSide 'scalemode set here
    
    bNoStats = True
    
    lCardType = CongregationMember.GetPublisherCardType(PersonID)
    
    'get all the print fields....
    With CongregationMember
    
    If PrintHeader Then
    
        msName = .FullNameFromDB_Official(PersonID)
        
        msAddress = .Address1(PersonID)
        If .Address2(PersonID) <> "" Then
            msAddress = msAddress & ", " & .Address2(PersonID)
        End If
        If .Address3(PersonID) <> "" Then
            msAddress = msAddress & ", " & .Address3(PersonID)
        End If
        If .Address4(PersonID) <> "" Then
            msAddress = msAddress & ", " & .Address4(PersonID)
        End If
        If .PostCode(PersonID) <> "" Then
            msAddress = msAddress & ", " & .PostCode(PersonID)
        End If
        
        Select Case lCardType
        Case 1
            msTelNo = .HomePhone(PersonID)
            msMobileNo = .MobilePhone(PersonID)
        Case Else
            msTelNo = .HomePhone(PersonID)
            If msTelNo = "" Then
                msTelNo = .MobilePhone(PersonID)
            End If
            msMobileNo = ""
        End Select
        
        Select Case lCardType
        Case 1
            msIsFemale = IIf(.IsFemale(PersonID), "x", "")
            msIsMale = IIf(.IsMale(PersonID), "x", "")
        Case Else
            msIsFemale = ""
            msIsMale = ""
        End Select
            
        tempvar = .ElderDate(PersonID, TempVar2)
        msIsElder = IIf(TempVar2, "x", "")
        
        tempvar = .ServantDate(PersonID, TempVar2)
        msIsMS = IIf(TempVar2, "x", "")
        
        msIsPio = IIf(.IsRegPio(PersonID, Now, msPioNo), "x", "")
    
        msAnointed = IIf(.IsAnointed(PersonID), "A", "OS")
        
        msBaptDate = IIf(.BaptismDate(PersonID) > 0, CStr(.BaptismDate(PersonID)), "")
        
        msDOB = IIf(.DateOfBirth(PersonID) > 0, CStr(.DateOfBirth(PersonID)), "")
        
        
        If Not CanDoThePrintHeader Then 'any fields too long for the card?
            PrintPublisherCard = False
            Exit Function
        Else
            bNoStats = False
        End If
        
    End If
    
    End With
    
    
    DefaultPageWidth = Printer.Width
    DefaultPageHeight = Printer.Height
    
    Printer.Width = 566.929 * PubCardPaperWidth
    Printer.Height = 566.929 * PubCardPaperHeight
            
    If PrintHeader Then
        Select Case lCardType
        Case 1
            PrintMale
        End Select
        
        PrintPioNumber
        
        Select Case lCardType
        Case 1
            PrintFemale
        End Select
        
        PrintPublisherName
        PrintAddress
        PrintTelNo
        
        Select Case lCardType
        Case 1
            PrintMobileNo
        End Select
        
        PrintDOB
        PrintElder
        PrintServant
        PrintRegPio
        PrintBaptismDate
        PrintAnointed
        bNoStats = False
    End If
    
    If PrintYear Then
        PrintServiceYear ServiceYear
        bNoStats = False
    End If
            
    If PrintStats Then
    
        PrintSQL = IIf(UnprintedOnly, " AND Printed = FALSE ", "")
    
        TheString = "SELECT NoBooks, " & _
                    "       NoBooklets, " & _
                    "       NoHours, " & _
                    "       NoMagazines, " & _
                    "       NoReturnVisits, " & _
                    "       NoStudies, " & _
                    "       NoTracts, " & _
                    "       Remarks, " & _
                    "       MinistryDoneInMonth, " & _
                    "       OrderForServiceYear, " & _
                    "       tblMinReports.ActualMinPeriod " & _
                    "FROM (tblMinReports INNER JOIN tblPubRecCardRowPrinted ON " & _
                    "(tblMinReports.ActualMinPeriod = tblPubRecCardRowPrinted.ActualMinPeriod) " & _
                    "AND (tblMinReports.PersonID = tblPubRecCardRowPrinted.PersonID)) " & _
                    "INNER JOIN tblMonthName ON tblMinReports.MinistryDoneInMonth = tblMonthName.MonthNum " & _
                    "WHERE tblMinReports.PersonID = " & PersonID & _
                   " AND MinistryDoneInYear = " & ServiceYear & _
                   PrintSQL & _
                   " ORDER BY OrderForServiceYear "
                   
        Set rstMinStats = CMSDB.OpenRecordset(TheString, dbOpenSnapshot)
        
        With rstMinStats
        
        If bNoStats Then
            bNoStats = .BOF
        End If
            
        Do Until .EOF Or .BOF
            msBooks = IIf(!NoBooks > 0, !NoBooks, "")
            msBrochures = IIf(!NoBooklets > 0, !NoBooklets, "")
            msHours = IIf(!NoHours > 0, !NoHours, "")
            msMags = IIf(!NoMagazines > 0, !NoMagazines, "")
            msRVs = IIf(!NoReturnVisits > 0, !NoReturnVisits, "")
            msStudies = IIf(!NoStudies > 0, !NoStudies, "")
            msRemarks = !Remarks
            PrintMonthStats !OrderForServiceYear
            SetRecCardRowPrintedFlag True, PersonID, CStr(!ActualMinPeriod)
            .MoveNext
        Loop
            
        frmCongStats.GetYearReport 'refresh printed column of grid
        
        End With
        
        If ShouldWePrintTotals(ServiceYear, 1) Then
        
            TheString = "SELECT SUM(NoBooks) as SumBooks, " & _
                        "       SUM(NoBooklets) as SumBooklets, " & _
                        "       SUM(NoHours) as SumHours, " & _
                        "       SUM(NoMagazines) as SumMags, " & _
                        "       SUM(NoReturnVisits) as SumRVs, " & _
                        "       SUM(NoStudies) as SumStudies, " & _
                        "       SUM(NoTracts) as SumTracts " & _
                        "FROM (tblMinReports INNER JOIN tblPubRecCardRowPrinted ON " & _
                        "(tblMinReports.ActualMinPeriod = tblPubRecCardRowPrinted.ActualMinPeriod) " & _
                        "AND (tblMinReports.PersonID = tblPubRecCardRowPrinted.PersonID)) " & _
                        "INNER JOIN tblMonthName ON tblMinReports.MinistryDoneInMonth = tblMonthName.MonthNum " & _
                        "WHERE tblMinReports.PersonID = " & PersonID & _
                       " AND MinistryDoneInYear = " & ServiceYear
                       
            Set rstMinStats = CMSDB.OpenRecordset(TheString, dbOpenSnapshot)
            
            With rstMinStats
            
            If Not .BOF Then
                msBooks = IIf(!SumBooks > 0, !SumBooks, "")
                msBrochures = IIf(!SumBooklets > 0, !SumBooklets, "")
                msHours = IIf(!SumHours > 0, !SumHours, "")
                msMags = IIf(!SumMags > 0, !SumMags, "")
                msRVs = IIf(!SumRVs > 0, !SumRVs, "")
                msStudies = IIf(!SumStudies > 0, !SumStudies, "")
                msRemarks = ""
                PrintMonthStats 13
                bNoStats = False
            End If

            End With
                       
        End If
        
        rstMinStats.Close
        
    End If
            
    Printer.EndDoc
    
    Printer.Width = DefaultPageWidth
    Printer.Height = DefaultPageHeight
        
    
    If bNoStats Then
        ShowMessage "Nothing to print", 1250, frmCongStats
    Else
        ShowMessage AddApostropheToPersonName(CongregationMember.NameWithMiddleInitial(PersonID)) & _
                " publisher card details have been sent to the default printer.", 2000, frmCongStats
    End If
        
    PrintPublisherCard = True
    
    Exit Function
ErrorTrap:
    MsgBox "There was a printing error. Please check printer and try again.", vbOKOnly + vbCritical, AppName
    PrintPublisherCard = False
    Exit Function
    
End Function

Private Sub SetUpPrintParameters(Optional PersonID As Long = 0, Optional CardSide As Long = 1)
Dim lCardType As Long, sParmVer As String, sCardSide As String
On Error GoTo ErrorTrap

    lCardType = CongregationMember.GetPublisherCardType(PersonID) 'if persoinID =0 then default card type returned

    Select Case lCardType
    Case 0
        sParmVer = ""
    Case 1
    
        sParmVer = "_4_05"
        
        Select Case CardSide
        Case 1
            sCardSide = ""
        Case 2
            sCardSide = "_Side2"
        End Select
        
    End Select
    
    Printer.ScaleMode = vbCentimeters
     
    Select Case lCardType
    Case 1
        PubCardMaleXPos = GlobalParms.GetValue("PubCardMaleXPos" & sParmVer, "NumFloat")
        PubCardMaleYPos = GlobalParms.GetValue("PubCardMaleYPos" & sParmVer, "NumFloat")
        PubCardFemaleXPos = GlobalParms.GetValue("PubCardFemaleXPos" & sParmVer, "NumFloat")
        PubCardFemaleYPos = GlobalParms.GetValue("PubCardFemaleYPos" & sParmVer, "NumFloat")
        PubCardMobileNoXPos = GlobalParms.GetValue("PubCardMobileNoXPos" & sParmVer, "NumFloat")
        PubCardMobileNoYPos = GlobalParms.GetValue("PubCardMobileNoYPos" & sParmVer, "NumFloat")
        PubCardMobileNoMaxWidth = GlobalParms.GetValue("PubCardMobileNoMaxWidth" & sParmVer, "NumFloat")
    Case Else
    End Select
    
    PubCardTweakX = GlobalParms.GetValue("PubCardTweakX" & sCardSide & sParmVer, "NumFloat")
    PubCardTweakY = GlobalParms.GetValue("PubCardTweakY" & sCardSide & sParmVer, "NumFloat")
     
    PubCardPublisherNameXPos = GlobalParms.GetValue("PubCardPublisherNameXPos" & sParmVer, "NumFloat")
    PubCardPublisherNameYPos = GlobalParms.GetValue("PubCardPublisherNameYPos" & sParmVer, "NumFloat")
    PubCardAddressXPos = GlobalParms.GetValue("PubCardAddressXPos" & sParmVer, "NumFloat")
    PubCardAddressYPos = GlobalParms.GetValue("PubCardAddressYPos" & sParmVer, "NumFloat")
    PubCardPublisherNameMaxWidth = GlobalParms.GetValue("PubCardPublisherNameMaxWidth" & sParmVer, "NumFloat")
    PubCardAddressMaxWidth = GlobalParms.GetValue("PubCardAddressMaxWidth" & sParmVer, "NumFloat")
    PubCardTelNoXPos = GlobalParms.GetValue("PubCardTelNoXPos" & sParmVer, "NumFloat")
    PubCardTelNoYPos = GlobalParms.GetValue("PubCardTelNoYPos" & sParmVer, "NumFloat")
    PubCardTelNoMaxWidth = GlobalParms.GetValue("PubCardTelNoMaxWidth" & sParmVer, "NumFloat")
    PubCardElderXPos = GlobalParms.GetValue("PubCardElderXPos" & sParmVer, "NumFloat")
    PubCardElderYPos = GlobalParms.GetValue("PubCardElderYPos" & sParmVer, "NumFloat")
    PubCardServantXPos = GlobalParms.GetValue("PubCardServantXPos" & sParmVer, "NumFloat")
    PubCardServantYPos = GlobalParms.GetValue("PubCardServantYPos" & sParmVer, "NumFloat")
    PubCardRegPioXPos = GlobalParms.GetValue("PubCardRegPioXPos" & sParmVer, "NumFloat")
    PubCardRegPioYPos = GlobalParms.GetValue("PubCardRegPioYPos" & sParmVer, "NumFloat")
    PubCardBaptDateXPos = GlobalParms.GetValue("PubCardBaptDateXPos" & sParmVer, "NumFloat")
    PubCardBaptDateYPos = GlobalParms.GetValue("PubCardBaptDateYPos" & sParmVer, "NumFloat")
    PubCardDOBXPos = GlobalParms.GetValue("PubCardDOBXPos" & sParmVer, "NumFloat")
    PubCardDOBYPos = GlobalParms.GetValue("PubCardDOBYPos" & sParmVer, "NumFloat")
    PubCardBaptDateMaxWidth = GlobalParms.GetValue("PubCardBaptDateMaxWidth" & sParmVer, "NumFloat")
    PubCardPioNoXPos = GlobalParms.GetValue("PubCardPioNoXPos" & sParmVer, "NumFloat")
    PubCardPioNoYPos = GlobalParms.GetValue("PubCardPioNoYPos" & sParmVer, "NumFloat")
    PubCardPioNoMaxWidth = GlobalParms.GetValue("PubCardPioNoMaxWidth" & sParmVer, "NumFloat")
    PubCardAnointedXPos = GlobalParms.GetValue("PubCardAnointedXPos" & sParmVer, "NumFloat")
    PubCardAnointedYPos = GlobalParms.GetValue("PubCardAnointedYPos" & sParmVer, "NumFloat")
    PubCardAnointedMaxWidth = GlobalParms.GetValue("PubCardAnointedMaxWidth" & sParmVer, "NumFloat")
    PubCardTopMargin = GlobalParms.GetValue("PubCardTopMargin" & sParmVer, "NumFloat")
    PubCardLeftMargin = GlobalParms.GetValue("PubCardLeftMargin" & sParmVer, "NumFloat")
    PubCardFontSize = GlobalParms.GetValue("PubCardFontSize" & sParmVer, "NumFloat")
    PubCardFontName = GlobalParms.GetValue("PubCardFontName" & sParmVer, "AlphaVal")
    PubCardPaperHeight = GlobalParms.GetValue("PubCardPaperHeight" & sParmVer, "NumFloat")
    PubCardPaperWidth = GlobalParms.GetValue("PubCardPaperWidth" & sParmVer, "NumFloat")
    PubCardServiceYearXPos = GlobalParms.GetValue("PubCardServiceYearXPos" & sParmVer, "NumFloat")
    PubCardServiceYearYPos = GlobalParms.GetValue("PubCardServiceYearYPos" & sParmVer, "NumFloat")
    PubCardBooksXPos = GlobalParms.GetValue("PubCardBooksXPos" & sParmVer, "NumFloat")
    PubCardBrochuresXPos = GlobalParms.GetValue("PubCardBrochuresXPos" & sParmVer, "NumFloat")
    PubCardHoursXPos = GlobalParms.GetValue("PubCardHoursXPos" & sParmVer, "NumFloat")
    PubCardMagsXPos = GlobalParms.GetValue("PubCardMagsXPos" & sParmVer, "NumFloat")
    PubCardRVsXPos = GlobalParms.GetValue("PubCardRVsXPos" & sParmVer, "NumFloat")
    PubCardStudiesXPos = GlobalParms.GetValue("PubCardStudiesXPos" & sParmVer, "NumFloat")
    PubCardRemarksXPos = GlobalParms.GetValue("PubCardRemarksXPos" & sParmVer, "NumFloat")
    PubCardSeptYPos = GlobalParms.GetValue("PubCardSeptYPos" & sParmVer, "NumFloat")
    PubCardOctYPos = GlobalParms.GetValue("PubCardOctYPos" & sParmVer, "NumFloat")
    PubCardNovYPos = GlobalParms.GetValue("PubCardNovYPos" & sParmVer, "NumFloat")
    PubCardDecYPos = GlobalParms.GetValue("PubCardDecYPos" & sParmVer, "NumFloat")
    PubCardJanYPos = GlobalParms.GetValue("PubCardJanYPos" & sParmVer, "NumFloat")
    PubCardFebYPos = GlobalParms.GetValue("PubCardFebYPos" & sParmVer, "NumFloat")
    PubCardMarYPos = GlobalParms.GetValue("PubCardMarYPos" & sParmVer, "NumFloat")
    PubCardAprYPos = GlobalParms.GetValue("PubCardAprYPos" & sParmVer, "NumFloat")
    PubCardMayYPos = GlobalParms.GetValue("PubCardMayYPos" & sParmVer, "NumFloat")
    PubCardJunYPos = GlobalParms.GetValue("PubCardJunYPos" & sParmVer, "NumFloat")
    PubCardJulYPos = GlobalParms.GetValue("PubCardJulYPos" & sParmVer, "NumFloat")
    PubCardAugYPos = GlobalParms.GetValue("PubCardAugYPos" & sParmVer, "NumFloat")
    PubCardTotalYPos = GlobalParms.GetValue("PubCardTotalYPos" & sParmVer, "NumFloat")
    PubCardRemarksMaxWidth = GlobalParms.GetValue("PubCardRemarksMaxWidth" & sParmVer, "NumFloat")

    Printer.Font.Name = PubCardFontName
    Printer.Font.Size = PubCardFontSize


    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintMonthStats(MonthNoForServiceYear)

On Error GoTo ErrorTrap

Dim YPos As Single

    Select Case MonthNoForServiceYear
    Case 1
        YPos = PubCardSeptYPos
    Case 2
        YPos = PubCardOctYPos
    Case 3
        YPos = PubCardNovYPos
    Case 4
        YPos = PubCardDecYPos
    Case 5
        YPos = PubCardJanYPos
    Case 6
        YPos = PubCardFebYPos
    Case 7
        YPos = PubCardMarYPos
    Case 8
        YPos = PubCardAprYPos
    Case 9
        YPos = PubCardMayYPos
    Case 10
        YPos = PubCardJunYPos
    Case 11
        YPos = PubCardJulYPos
    Case 12
        YPos = PubCardAugYPos
    Case 13
        YPos = PubCardTotalYPos
    End Select

    With Printer
    
    'Must put .CurrentY statement prior to each Print, even if identical (!)
    
    .CurrentY = YPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters) + PubCardTweakY
    .CurrentX = PubCardBooksXPos + PubCardTweakX
    Printer.Print msBooks
    
    .CurrentY = YPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters) + PubCardTweakY
    .CurrentX = PubCardBrochuresXPos + PubCardTweakX
    Printer.Print msBrochures
    
    .CurrentY = YPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters) + PubCardTweakY
    .CurrentX = PubCardHoursXPos + PubCardTweakX
    Printer.Print msHours
    
    .CurrentY = YPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters) + PubCardTweakY
    .CurrentX = PubCardMagsXPos + PubCardTweakX
    Printer.Print msMags
    
    .CurrentY = YPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters) + PubCardTweakY
    .CurrentX = PubCardRVsXPos + PubCardTweakX
    Printer.Print msRVs
    
    .CurrentY = YPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters) + PubCardTweakY
    .CurrentX = PubCardStudiesXPos + PubCardTweakX
    Printer.Print msStudies
    
    .CurrentY = YPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters) + PubCardTweakY
    .CurrentX = PubCardRemarksXPos + PubCardTweakX
    Printer.Print TruncateTextToFit(msRemarks, PubCardRemarksMaxWidth, _
                                    PubCardFontName, PubCardFontSize)
    
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintPublisherName()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = PubCardPublisherNameXPos
    .CurrentY = PubCardPublisherNameYPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + PubCardTweakX
    .CurrentY = .CurrentY + PubCardTweakY
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msName
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintServiceYear(ServiceYear As String)

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = PubCardServiceYearXPos
    .CurrentY = PubCardServiceYearYPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + PubCardTweakX
    .CurrentY = .CurrentY + PubCardTweakY
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print ServiceYear
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PrintAddress()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = PubCardAddressXPos
    .CurrentY = PubCardAddressYPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + PubCardTweakX
    .CurrentY = .CurrentY + PubCardTweakY
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msAddress
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintTelNo()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = PubCardTelNoXPos
    .CurrentY = PubCardTelNoYPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + PubCardTweakX
    .CurrentY = .CurrentY + PubCardTweakY
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msTelNo
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintMobileNo()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = PubCardMobileNoXPos
    .CurrentY = PubCardMobileNoYPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + PubCardTweakX
    .CurrentY = .CurrentY + PubCardTweakY
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msMobileNo
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintElder()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = PubCardElderXPos
    .CurrentY = PubCardElderYPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + PubCardTweakX
    .CurrentY = .CurrentY + PubCardTweakY
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msIsElder
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintMale()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = PubCardMaleXPos
    .CurrentY = PubCardMaleYPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + PubCardTweakX
    .CurrentY = .CurrentY + PubCardTweakY
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msIsMale
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintFemale()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = PubCardFemaleXPos
    .CurrentY = PubCardFemaleYPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + PubCardTweakX
    .CurrentY = .CurrentY + PubCardTweakY
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msIsFemale
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintServant()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = PubCardServantXPos
    .CurrentY = PubCardServantYPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + PubCardTweakX
    .CurrentY = .CurrentY + PubCardTweakY
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msIsMS
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintRegPio()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = PubCardRegPioXPos
    .CurrentY = PubCardRegPioYPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + PubCardTweakX
    .CurrentY = .CurrentY + PubCardTweakY
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msIsPio
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintBaptismDate()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = PubCardBaptDateXPos
    .CurrentY = PubCardBaptDateYPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + PubCardTweakX
    .CurrentY = .CurrentY + PubCardTweakY
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msBaptDate
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintDOB()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = PubCardDOBXPos
    .CurrentY = PubCardDOBYPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + PubCardTweakX
    .CurrentY = .CurrentY + PubCardTweakY
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msDOB
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintPioNumber()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = PubCardPioNoXPos
    .CurrentY = PubCardPioNoYPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + PubCardTweakX
    .CurrentY = .CurrentY + PubCardTweakY
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msPioNo
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PrintAnointed()

On Error GoTo ErrorTrap

    With Printer
    .CurrentX = PubCardAnointedXPos
    .CurrentY = PubCardAnointedYPos - .ScaleY(PubCardFontSize, vbPoints, vbCentimeters)
    .CurrentX = .CurrentX + PubCardTweakX
    .CurrentY = .CurrentY + PubCardTweakY
    
    '
    'For some reason, simply putting '.Print' doesn't work...
    '
    Printer.Print msAnointed
    End With

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Public Sub SetRecCardRowPrintedFlag(Printed As Boolean, _
                                     PersonID As Long, _
                                     ActualMinPeriod As String)

On Error GoTo ErrorTrap

    CMSDB.Execute "UPDATE tblPubRecCardRowPrinted " & _
                    "SET Printed = " & Printed & _
                    " WHERE PersonID = " & PersonID & _
                    " AND ActualMinPeriod = #" & Format(ActualMinPeriod, "mm/dd/yyyy") & "#"
    
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Public Sub SetCongRecCardRowPrintedFlag(Printed As Boolean, _
                                        ServiceYear As Long, _
                                        MinMonth As Long, _
                                        MinType As Long)

On Error GoTo ErrorTrap
Dim str As String, rs As Recordset, i As Long
          
    str = "SELECT MinMonth, MinServiceYear, MinType " & _
          "FROM tblCongMinCardRowPrinted " & _
          "WHERE MinServiceYear = " & ServiceYear & _
          " AND MinType = " & MinType & _
          " ORDER BY MinMonth "
          
    Set rs = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    With rs
    .FindFirst "MinMonth = " & MinMonth
    If Printed Then
        If .NoMatch Then
            .AddNew
            !MinServiceYear = ServiceYear
            !MinType = MinType
            !MinMonth = MinMonth
            .Update
        End If
    Else
        If Not .NoMatch Then
            .Delete
        End If
    End If
    
    End With
    
    rs.Close
    Set rs = Nothing
    
    Exit Sub
ErrorTrap:
    Call EndProgram

    
End Sub
Private Function CongRecCardRowPrinted(ServiceYear As Long, _
                                        MinMonth As Long, _
                                        MinType As Long) As Boolean

On Error GoTo ErrorTrap
Dim str As String, rs As Recordset, i As Long
          
    str = "SELECT MinMonth, MinServiceYear, MinType " & _
          "FROM tblCongMinCardRowPrinted " & _
          "WHERE MinServiceYear = " & ServiceYear & _
          " AND MinType = " & MinType & _
          " ORDER BY MinMonth "
          
    Set rs = CMSDB.OpenRecordset(str, dbOpenDynaset)
    
    With rs
    .FindFirst "MinMonth = " & MinMonth
    If .NoMatch Then
        CongRecCardRowPrinted = False
    Else
        CongRecCardRowPrinted = True
    End If
    
    End With
    
    rs.Close
    Set rs = Nothing
    
    Exit Function
ErrorTrap:
    Call EndProgram

    
End Function
Private Function ShouldWePrintTotals(ServiceYear As String, Mode As Long) As Boolean

On Error GoTo ErrorTrap
Dim str As String, CurrentServiceYear As String
          
    If Mode = 1 Then
        If frmCongStats.chkTotals = vbUnchecked Then
            ShouldWePrintTotals = False
            Exit Function
        End If
    Else
        If frmPrintCongMinCard.chkTotals = vbUnchecked Then
            ShouldWePrintTotals = False
            Exit Function
        End If
    End If
        
        
    CurrentServiceYear = year(ConvertNormalDateToServiceDate(Now))
    
    If CurrentServiceYear > ServiceYear Then
        ShouldWePrintTotals = True
        Exit Function
    End If
        
    If Mode = 1 Then
        ShowMessage "Service Year totals will not be printed", 1500, frmCongStats, , vbRed
    Else
        ShowMessage "Service Year totals will not be printed", 1500, frmPrintCongMinCard, , vbRed
    End If
        
    ShouldWePrintTotals = False
    
    Exit Function
ErrorTrap:
    Call EndProgram

    
End Function

Private Function CanDoThePrintHeader() As Boolean

On Error GoTo ErrorTrap

    If Printer.TextWidth(msName) > PubCardPublisherNameMaxWidth Then
        MsgBox "Publisher name too long for card", vbExclamation + vbOKOnly, AppName
        CanDoThePrintHeader = False
        Exit Function
    End If
    
    If Printer.TextWidth(msAddress) > PubCardAddressMaxWidth Then
        MsgBox "Publisher address too long for card", vbExclamation + vbOKOnly, AppName
        CanDoThePrintHeader = False
        Exit Function
    End If
    
    If Printer.TextWidth(msTelNo) > PubCardTelNoMaxWidth Then
        MsgBox "Home telephone number too long for card", vbExclamation + vbOKOnly, AppName
        CanDoThePrintHeader = False
        Exit Function
    End If
    
    If Printer.TextWidth(msPioNo) > PubCardPioNoMaxWidth Then
        MsgBox "Pioneer Number too long for card", vbExclamation + vbOKOnly, AppName
        CanDoThePrintHeader = False
        Exit Function
    End If
    
    If Printer.TextWidth(msMobileNo) > PubCardMobileNoMaxWidth Then
        MsgBox "Mobile Number too long for card", vbExclamation + vbOKOnly, AppName
        CanDoThePrintHeader = False
        Exit Function
    End If
    
    CanDoThePrintHeader = True
    
    Exit Function
ErrorTrap:
    EndProgram
    
End Function
Private Function AddHourFractions(MinPeriod As String, HoursSoFar As Double) As Double

On Error GoTo ErrorTrap

Dim AnAmount As Double, TheString As String, rstRecSet As Recordset

    If Len(MinPeriod) = 10 Then
    
        TheString = "SELECT SUM(NoHours) AS TheSum " & _
                    "FROM tblMinReports " & _
                    "WHERE tblMinReports.ActualMinPeriod < " & GetDateStringForSQLWhere(MinPeriod)
                    
        
    Else
    
        TheString = "SELECT SUM(NoHours) AS TheSum " & _
                    "FROM tblMinReports " & _
                    "WHERE tblMinReports.MinistryDoneInYear < " & MinPeriod
                    
            
    End If
    
    Set rstRecSet = CMSDB.OpenRecordset(TheString, dbOpenSnapshot)
        
    With rstRecSet
    
    If Not .BOF Then
        If Not IsNull(!TheSum) Then
            AnAmount = GetFractionPart(!TheSum)
        Else
            AnAmount = 0
        End If
    Else
        AnAmount = 0
    End If
    
    AddHourFractions = Fix(HoursSoFar + AnAmount)
               
    End With
    
    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Public Function PrintCongMinCard(ByVal MinType As Long, _
                                   ByVal PrintHeader As Boolean, _
                                   ByVal PrintYear As Boolean, _
                                   ByVal PrintStats As Boolean, _
                                   ByVal UnprintedOnly As Boolean, _
                                   ByVal PrintTotals As Boolean, _
                                   ByVal ServiceYear As String, _
                                   ByVal bIncludeZeroHours As Boolean, _
                                   ByVal CardSide As Long) As Boolean

On Error GoTo ErrorTrap
Dim rstMinStats As Recordset, TheString As String, SQL1 As String
Dim SQL2 As String, bNoStats As Boolean, ZeroHourSQL As String, lHours As Long
Dim dHours As Double

    '
    'Since we're using same card as that for individual publishers, same
    ' fields are used here
    '
    
    bNoStats = True
    
    SetUpPrintParameters , CardSide  'scalemode set here
    
    'get all the print fields....
    
    DefaultPageWidth = Printer.Width
    DefaultPageHeight = Printer.Height
    
    Printer.Width = 566.929 * PubCardPaperWidth
    Printer.Height = 566.929 * PubCardPaperHeight
            
    Select Case MinType
    Case IsPublisher
        msName = "Publishers"
        SQL1 = ""
        SQL2 = " AND A.PersonID NOT IN" & _
               "   (SELECT PersonID FROM tblAuxPioDates WHERE A.ActualMinPeriod BETWEEN StartDate AND EndDate) " & _
               " AND A.PersonID NOT IN" & _
               "   (SELECT PersonID FROM tblRegPioDates WHERE A.ActualMinPeriod BETWEEN StartDate AND EndDate) " & _
               " AND A.PersonID NOT IN" & _
               "   (SELECT PersonID FROM tblSpecPioDates WHERE A.ActualMinPeriod BETWEEN StartDate AND EndDate) "
               
    Case IsAuxPio
        msName = "Auxilliary Pioneers"
        SQL1 = " INNER JOIN tblAuxPioDates AS B ON A.PersonID = B.PersonID "
        SQL2 = "AND A.ActualMinPeriod BETWEEN StartDate AND EndDate"
    Case IsRegPio
        msName = "Regular Pioneers"
        SQL1 = " INNER JOIN tblRegPioDates AS B ON A.PersonID = B.PersonID "
        SQL2 = "AND A.ActualMinPeriod BETWEEN StartDate AND EndDate"
    Case IsSpecPio
        msName = "Special Pioneers"
        SQL1 = " INNER JOIN tblSpecPioDates AS B ON A.PersonID = B.PersonID "
        SQL2 = "AND A.ActualMinPeriod BETWEEN StartDate AND EndDate"
    End Select
    
    If PrintHeader Then
        PrintPublisherName
        bNoStats = False
    End If
    
    If PrintYear Then
        PrintServiceYear ServiceYear
        bNoStats = False
    End If
    
    If PrintStats Then
    
        'Include pubs where Hours are zero?
        If bIncludeZeroHours Then
            ZeroHourSQL = " "
        Else
            ZeroHourSQL = " AND NoHours > 0 "
        End If
        
        TheString = "SELECT SUM(NoBooks) as SumBooks, " & _
                    "       SUM(NoBooklets) as SumBooklets, " & _
                    "       SUM(NoHours) as SumHours, " & _
                    "       SUM(NoMagazines) as SumMags, " & _
                    "       SUM(NoReturnVisits) as SumRVs, " & _
                    "       SUM(NoStudies) as SumStudies, " & _
                    "       SUM(NoTracts) as SumTracts, " & _
                    "       COUNT(A.PersonID) as CountPubs, " & _
                    "       MinistryDoneInMonth, " & _
                    "       OrderForServiceYear, " & _
                    "       A.ActualMinPeriod AS ActMinPer " & _
                    "FROM (tblMinReports AS A INNER JOIN tblMonthName ON " & _
                    "     A.MinistryDoneInMonth = tblMonthName.MonthNum) " & _
                    SQL1 & _
                    " WHERE MinistryDoneInYear = " & ServiceYear & _
                    SQL2 & ZeroHourSQL & _
                   " GROUP BY MinistryDoneInMonth, OrderForServiceYear, A.ActualMinPeriod " & _
                   " ORDER BY OrderForServiceYear "

        Set rstMinStats = CMSDB.OpenRecordset(TheString, dbOpenSnapshot)
        
        With rstMinStats
        
        If bNoStats = True Then
            bNoStats = .BOF
        End If
        
        Do Until .EOF Or .BOF
            msBooks = IIf(HandleNull(!SumBooks) > 0, !SumBooks, "")
            msBrochures = IIf(HandleNull(!SumBooklets) > 0, !SumBooklets, "")
            
            lHours = AddHourFractions(!ActMinPer, HandleNull(!SumHours))
            msHours = lHours
            
            msMags = IIf(HandleNull(!SumMags) > 0, !SumMags, "")
            msRVs = IIf(HandleNull(!SumRVs) > 0, !SumRVs, "")
            msStudies = IIf(HandleNull(!SumStudies) > 0, !SumStudies, "")
            msRemarks = IIf(HandleNull(!CountPubs) > 0, !CountPubs, "")
            
            If UnprintedOnly Then
                If Not CongRecCardRowPrinted(CLng(ServiceYear), !MinistryDoneInMonth, MinType) Then
                    PrintMonthStats !OrderForServiceYear
                    SetCongRecCardRowPrintedFlag True, CLng(ServiceYear), _
                                                    !MinistryDoneInMonth, MinType
                End If
            Else
                PrintMonthStats !OrderForServiceYear
                SetCongRecCardRowPrintedFlag True, CLng(ServiceYear), _
                                                !MinistryDoneInMonth, MinType
            End If
                
            .MoveNext
        Loop
            
        frmPrintCongMinCard.FillGrid 'refresh printed column of grid
        
        End With
        
        If ShouldWePrintTotals(ServiceYear, 2) Then
        
            TheString = "SELECT SUM(NoBooks) as SumBooks, " & _
                        "       SUM(NoBooklets) as SumBooklets, " & _
                        "       SUM(NoHours) as SumHours, " & _
                        "       SUM(NoMagazines) as SumMags, " & _
                        "       SUM(NoReturnVisits) as SumRVs, " & _
                        "       SUM(NoStudies) as SumStudies, " & _
                        "       SUM(NoTracts) as SumTracts " & _
                        "FROM tblMinReports AS A " & _
                        SQL1 & _
                        "WHERE MinistryDoneInYear = " & ServiceYear & _
                        SQL2
                       
            Set rstMinStats = CMSDB.OpenRecordset(TheString, dbOpenSnapshot)
            
            With rstMinStats
            
            msBooks = IIf(HandleNull(!SumBooks) > 0, !SumBooks, "")
            msBrochures = IIf(HandleNull(!SumBooklets) > 0, !SumBooklets, "")
            
            lHours = AddHourFractions(ServiceYear, HandleNull(!SumHours))
            msHours = lHours
            
            msMags = IIf(HandleNull(!SumMags > 0), !SumMags, "")
            msRVs = IIf(HandleNull(!SumRVs > 0), !SumRVs, "")
            msStudies = IIf(HandleNull(!SumStudies > 0), !SumStudies, "")
                        
            End With
            
            'distinct count of all publishers...
            TheString = "SELECT COUNT(PersonID) AS CountPubs FROM " & _
                "(SELECT DISTINCT A.PersonID  " & _
                    "FROM tblMinReports AS A " & _
                    SQL1 & _
                    "WHERE MinistryDoneInYear = " & ServiceYear & _
                    SQL2 & ZeroHourSQL & ")"
                       
            Set rstMinStats = CMSDB.OpenRecordset(TheString, dbOpenSnapshot)
            
            With rstMinStats
            msRemarks = IIf(HandleNull(!CountPubs > 0), !CountPubs, "")
            End With
                       
            PrintMonthStats 13
            
            bNoStats = False
            
        End If
        
        rstMinStats.Close
        
    End If
            
    Printer.EndDoc
    
    Printer.Width = DefaultPageWidth
    Printer.Height = DefaultPageHeight
        
    
    If bNoStats Then
        ShowMessage "Nothing to print.", 2000, frmPrintCongMinCard
    Else
        ShowMessage "The congregation record card details have been sent to the default printer.", 1750, frmPrintCongMinCard
    End If
        
    PrintCongMinCard = True
    
    Exit Function
ErrorTrap:
    MsgBox "There was a printing error. Please check printer and try again.", vbOKOnly + vbCritical, AppName
    PrintCongMinCard = False
    Exit Function
    
End Function

Public Function GetPubCardInfo(CardTypeID As Long) As String
                                     
Dim str As String, rs As Recordset

On Error GoTo ErrorTrap

    str = "SELECT CardSideInfo FROM tblPubCardTypes " & _
            "WHERE CardTypeID = " & CardTypeID
                
    Set rs = CMSDB.OpenRecordset(str, dbOpenForwardOnly)
    
    If rs.BOF Then
        GetPubCardInfo = ""
    Else
        GetPubCardInfo = HandleNull(rs!CardSideInfo, "")
    End If

    Exit Function
ErrorTrap:
    EndProgram
    
End Function


