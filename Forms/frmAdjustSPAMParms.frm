VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmAdjustSPAMParms 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " C.M.S. Sound & Platform Rota Parameters"
   ClientHeight    =   7785
   ClientLeft      =   330
   ClientTop       =   -15
   ClientWidth     =   4635
   Icon            =   "frmAdjustSPAMParms.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Max Consecutive"
      Height          =   1215
      Left            =   195
      TabIndex        =   18
      Top             =   5535
      Width           =   4245
      Begin VB.ComboBox Combo33 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmAdjustSPAMParms.frx":0442
         Left            =   3345
         List            =   "frmAdjustSPAMParms.frx":0455
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   645
         Width           =   712
      End
      Begin VB.ComboBox Combo32 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmAdjustSPAMParms.frx":0468
         Left            =   2295
         List            =   "frmAdjustSPAMParms.frx":047B
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   645
         Width           =   712
      End
      Begin VB.ComboBox Combo31 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmAdjustSPAMParms.frx":048E
         Left            =   1215
         List            =   "frmAdjustSPAMParms.frx":04A1
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   645
         Width           =   697
      End
      Begin VB.ComboBox Combo30 
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "frmAdjustSPAMParms.frx":04B4
         Left            =   120
         List            =   "frmAdjustSPAMParms.frx":04C7
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   645
         Width           =   712
      End
      Begin VB.Label Label48 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Platform"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   3330
         TabIndex        =   26
         Top             =   330
         Width           =   705
      End
      Begin VB.Label Label46 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sound"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   2265
         TabIndex        =   25
         Top             =   330
         Width           =   735
      End
      Begin VB.Label Label44 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Roving"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   1200
         TabIndex        =   24
         Top             =   330
         Width           =   735
      End
      Begin VB.Label Label41 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Attending"
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   105
         TabIndex        =   23
         Top             =   330
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   511
      Left            =   3330
      TabIndex        =   0
      Top             =   6960
      Width           =   1110
   End
   Begin VB.Frame Frame1 
      Caption         =   "Weightings"
      Height          =   5040
      Left            =   195
      TabIndex        =   1
      Top             =   390
      Width           =   4245
      Begin VB.CommandButton cmdAdjSPAMWtgPerPerson 
         Caption         =   "Adjust Person Weightings..."
         Height          =   511
         Left            =   1605
         TabIndex        =   39
         Top             =   645
         Width           =   1140
      End
      Begin VB.CommandButton cmdAdjSPAMWtgPerTask 
         Caption         =   "Adjust Task Weightings..."
         Height          =   511
         Left            =   2790
         TabIndex        =   38
         Top             =   645
         Width           =   1140
      End
      Begin VB.TextBox txtPTTrigger 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   316
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   35
         Top             =   3555
         Width           =   567
      End
      Begin ComctlLib.Slider sldrRespWtg 
         Height          =   165
         Left            =   1545
         TabIndex        =   27
         Top             =   375
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   291
         _Version        =   327682
         LargeChange     =   100
         Max             =   1000
         TickStyle       =   3
         TickFrequency   =   1000
      End
      Begin VB.TextBox txtInfirmity 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   316
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   2490
         Width           =   567
      End
      Begin VB.TextBox txtWkWtDiff 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   316
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   4545
         Visible         =   0   'False
         Width           =   567
      End
      Begin VB.TextBox txtMaxWkWtg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   316
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   4200
         Width           =   567
      End
      Begin VB.TextBox txtOverallPers 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   316
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   2955
         Width           =   567
      End
      Begin VB.TextBox txtLoneParent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   316
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2160
         Width           =   567
      End
      Begin VB.TextBox txtMarriage 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   316
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1815
         Width           =   567
      End
      Begin VB.TextBox txtParent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   316
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1470
         Width           =   567
      End
      Begin VB.TextBox txtRespWtg 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   316
         Left            =   3360
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   285
         Width           =   567
      End
      Begin ComctlLib.Slider sldrParent 
         Height          =   165
         Left            =   1545
         TabIndex        =   28
         Top             =   1590
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   291
         _Version        =   327682
         LargeChange     =   100
         Max             =   1000
         TickStyle       =   3
         TickFrequency   =   1000
      End
      Begin ComctlLib.Slider sldrMarriage 
         Height          =   165
         Left            =   1545
         TabIndex        =   29
         Top             =   1935
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   291
         _Version        =   327682
         LargeChange     =   100
         Max             =   1000
         TickStyle       =   3
         TickFrequency   =   1000
      End
      Begin ComctlLib.Slider sldrLoneParent 
         Height          =   165
         Left            =   1545
         TabIndex        =   30
         Top             =   2295
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   291
         _Version        =   327682
         LargeChange     =   100
         Max             =   1000
         TickStyle       =   3
         TickFrequency   =   1000
      End
      Begin ComctlLib.Slider sldrInfirmity 
         Height          =   165
         Left            =   1545
         TabIndex        =   31
         Top             =   2610
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   291
         _Version        =   327682
         LargeChange     =   100
         Max             =   1000
         TickStyle       =   3
         TickFrequency   =   1000
      End
      Begin ComctlLib.Slider sldrOverallPers 
         Height          =   165
         Left            =   1545
         TabIndex        =   32
         Top             =   3090
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   291
         _Version        =   327682
         LargeChange     =   100
         Max             =   1000
         TickStyle       =   3
         TickFrequency   =   1000
      End
      Begin ComctlLib.Slider sldrMaxWkWtg 
         Height          =   165
         Left            =   1545
         TabIndex        =   33
         Top             =   4320
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   291
         _Version        =   327682
         LargeChange     =   100
         Max             =   1000
         TickStyle       =   3
         TickFrequency   =   1000
      End
      Begin ComctlLib.Slider sldrWkWtDiff 
         Height          =   165
         Left            =   1545
         TabIndex        =   34
         Top             =   4650
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   291
         _Version        =   327682
         LargeChange     =   100
         Max             =   1000
         TickStyle       =   3
         TickFrequency   =   1000
      End
      Begin ComctlLib.Slider sldrPTTrigger 
         Height          =   165
         Left            =   1545
         TabIndex        =   36
         Top             =   3690
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   291
         _Version        =   327682
         LargeChange     =   100
         Max             =   1000
         TickStyle       =   3
         TickFrequency   =   1000
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Part Time Trigger"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   165
         TabIndex        =   37
         Top             =   3615
         Width           =   1425
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Infirmity Level"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   9
         Top             =   2550
         Width           =   1230
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Rota Repetition"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   165
         TabIndex        =   8
         Top             =   4605
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Lone Parent"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   7
         Top             =   2220
         Width           =   1155
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Last on Rota"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   165
         TabIndex        =   6
         Top             =   4260
         Width           =   1155
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Overall Personal"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   5
         Top             =   3015
         Width           =   1230
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Marriage"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   4
         Top             =   1875
         Width           =   1155
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Parent"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   3
         Top             =   1515
         Width           =   1155
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Responsibilities"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   150
         TabIndex        =   2
         Top             =   345
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmAdjustSPAMParms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SPAM_Parms As clsApplicationConstants
Dim mbIgnore As Boolean

'
'Use these values to calibrate the sliders. Reduce values to increase their effect...
'
Const InfirmityRange  As Single = 1000
Const RespWtgRange  As Single = 500
Const LoneParentRange As Single = 500
Const MarriageRange As Single = 3500
Const ParentRange As Single = 4500
Const OverallRange As Single = 3500
Const MaxWtgRange As Single = 0.1
Const WkWtDiffRange As Single = 200
Const PartTimeTriggerRange As Single = 1000


Private Sub cmdAdjSPAMWtgPerPerson_Click()
On Error GoTo ErrorTrap

    frmIndividualSPAMWeightings.Show vbModal, Me

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdAdjSPAMWtgPerTask_Click()
On Error GoTo ErrorTrap

    frmAdjustSPAMWtgForTasks.Show vbModal, Me

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdClose_Click()

On Error GoTo ErrorTrap

        Unload Me

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub sldrInfirmity_Change()

On Error GoTo ErrorTrap

    If mbIgnore Then Exit Sub

    txtInfirmity = Me!sldrInfirmity.value
    SPAM_Parms.Save "InfirmityLevelWtg", "NumFloat", CDbl(Me!sldrInfirmity.value / InfirmityRange)
    ShowMessage "Change saved", 500, Me

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub sldrPTTrigger_Change()

On Error GoTo ErrorTrap

    If mbIgnore Then Exit Sub

    txtPTTrigger = Me!sldrPTTrigger.value
    SPAM_Parms.Save "ThresholdForPartTimeSPAM", "NumFloat", CDbl(Me!sldrPTTrigger.value / PartTimeTriggerRange)
    ShowMessage "Change saved", 500, Me

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub sldrRespWtg_Change()

On Error GoTo ErrorTrap

    If mbIgnore Then Exit Sub

    txtRespWtg = Me!sldrRespWtg.value
    SPAM_Parms.Save "RespCoeff", "NumFloat", CDbl(Me!sldrRespWtg.value / RespWtgRange)
    ShowMessage "Change saved", 500, Me

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub sldrLoneParent_Change()

On Error GoTo ErrorTrap

    If mbIgnore Then Exit Sub

    txtLoneParent = Me!sldrLoneParent.value
    SPAM_Parms.Save "LoneParentFactor", "NumFloat", CDbl(Me!sldrLoneParent.value / LoneParentRange)
    'CalculateNewPartTimerThreshold
    ShowMessage "Change saved", 500, Me

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub sldrMarriage_Change()

On Error GoTo ErrorTrap

    If mbIgnore Then Exit Sub

    txtMarriage = Me!sldrMarriage.value
    SPAM_Parms.Save "Marriage_Wtg", "NumFloat", CDbl(Me!sldrMarriage.value / MarriageRange)
    ShowMessage "Change saved", 500, Me
    'CalculateNewPartTimerThreshold

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub sldrParent_Change()

On Error GoTo ErrorTrap

    If mbIgnore Then Exit Sub

    txtParent = Me!sldrParent.value
    SPAM_Parms.Save "ParentCoeff", "NumFloat", CDbl(Me!sldrParent.value / ParentRange)
    ShowMessage "Change saved", 500, Me
    'CalculateNewPartTimerThreshold

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub sldrOverallPers_Change()

On Error GoTo ErrorTrap

    If mbIgnore Then Exit Sub

    txtOverallPers = Me!sldrOverallPers.value
    SPAM_Parms.Save "PersonalWeightCoeff", "NumFloat", CDbl(Me!sldrOverallPers.value / OverallRange)
    ShowMessage "Change saved", 500, Me
    'CalculateNewPartTimerThreshold

    Exit Sub
ErrorTrap:
    EndProgram
End Sub
    
Private Sub sldrMaxWkWtg_Change()

On Error GoTo ErrorTrap

    If mbIgnore Then Exit Sub

    txtMaxWkWtg = Me!sldrMaxWkWtg.value
    SPAM_Parms.Save "MaxWkWting", "NumFloat", CDbl(Me!sldrMaxWkWtg.value / MaxWtgRange)
    ShowMessage "Change saved", 500, Me

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub sldrWkWtDiff_Change()

On Error GoTo ErrorTrap

    If mbIgnore Then Exit Sub

    txtWkWtDiff = Me!sldrWkWtDiff.value
    SPAM_Parms.Save "WkWtingDiff", "NumFloat", CDbl(Me!sldrWkWtDiff.value / WkWtDiffRange)
    ShowMessage "Change saved", 500, Me

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub



Private Sub Form_Load()

On Error GoTo ErrorTrap

'Set up new instance of class; Set initial slider values

    mbIgnore = True
    
    Set SPAM_Parms = New clsApplicationConstants
    
    sldrLoneParent.value = LoneParentRange * SPAM_Parms.GetValue("LoneParentFactor", "NumFloat")
    sldrMarriage.value = MarriageRange * SPAM_Parms.GetValue("Marriage_Wtg", "NumFloat")
    sldrRespWtg.value = RespWtgRange * SPAM_Parms.GetValue("RespCoeff", "NumFloat")
    sldrParent.value = ParentRange * SPAM_Parms.GetValue("ParentCoeff", "NumFloat")
    sldrOverallPers.value = OverallRange * SPAM_Parms.GetValue("PersonalWeightCoeff", "NumFloat")
    sldrMaxWkWtg.value = MaxWtgRange * SPAM_Parms.GetValue("MaxWkWting", "NumFloat")
    sldrWkWtDiff.value = WkWtDiffRange * SPAM_Parms.GetValue("WkWtingDiff", "NumFloat")
    sldrInfirmity.value = InfirmityRange * SPAM_Parms.GetValue("InfirmityLevelWtg", "NumFloat")
    sldrPTTrigger.value = PartTimeTriggerRange * SPAM_Parms.GetValue("ThresholdForPartTimeSPAM", "NumFloat")
    
    txtLoneParent.text = LoneParentRange * SPAM_Parms.GetValue("LoneParentFactor", "NumFloat")
    txtMarriage.text = MarriageRange * SPAM_Parms.GetValue("Marriage_Wtg", "NumFloat")
    txtRespWtg.text = RespWtgRange * SPAM_Parms.GetValue("RespCoeff", "NumFloat")
    txtParent.text = ParentRange * SPAM_Parms.GetValue("ParentCoeff", "NumFloat")
    txtOverallPers.text = OverallRange * SPAM_Parms.GetValue("PersonalWeightCoeff", "NumFloat")
    txtMaxWkWtg.text = MaxWtgRange * SPAM_Parms.GetValue("MaxWkWting", "NumFloat")
    txtWkWtDiff.text = WkWtDiffRange * SPAM_Parms.GetValue("WkWtingDiff", "NumFloat")
    txtInfirmity.text = InfirmityRange * SPAM_Parms.GetValue("InfirmityLevelWtg", "NumFloat")
    txtPTTrigger.text = PartTimeTriggerRange * SPAM_Parms.GetValue("ThresholdForPartTimeSPAM", "NumFloat")
    
    Combo30 = SPAM_Parms.GetValue("MaxConsecutiveAttendant", "NumVal")
    Combo31 = SPAM_Parms.GetValue("MaxConsecutiveRovingMic", "NumVal")
    Combo32 = SPAM_Parms.GetValue("MaxConsecutiveSound", "NumVal")
    Combo33 = SPAM_Parms.GetValue("MaxConsecutivePlatform", "NumVal")
    
    mbIgnore = False
    

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub CalculateNewPartTimerThreshold()

On Error GoTo ErrorTrap

Dim a As Double, b As Double, c As Double, d As Double

    If mbIgnore Then Exit Sub

'
'Adjust threshold for part-timer as the personal settings are adjusted
'

    'a = SPAM_Parms.GetValue("LoneParentFactor", "NumFloat")
    'b = SPAM_Parms.GetValue("Marriage_Wtg", "NumFloat")
    'c = SPAM_Parms.GetValue("ParentCoeff", "NumFloat")
    'd = SPAM_Parms.GetValue("PersonalWeightCoeff", "NumFloat")
    
    'SPAM_Parms.Save "ThresholdForPartTimeSPAM", "NumFloat", (a + b + c) * d * 4
    SPAM_Parms.Save "ThresholdForPartTimeSPAM", "NumFloat", 0.35


    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub Combo30_Click()

On Error GoTo ErrorTrap

    If mbIgnore Then Exit Sub

    GlobalParms.Save "MaxConsecutiveAttendant", "NumVal", CInt(Me!Combo30)
    GlobalParms.Save "SPAMOptionsChanged", "TrueFalse", True
    ShowMessage "Change saved", 500, Me

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub Combo31_Click()

On Error GoTo ErrorTrap

    If mbIgnore Then Exit Sub

    GlobalParms.Save "MaxConsecutiveRovingMic", "NumVal", CInt(Me!Combo31)
    GlobalParms.Save "SPAMOptionsChanged", "TrueFalse", True
    ShowMessage "Change saved", 500, Me

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub Combo32_Click()

On Error GoTo ErrorTrap

    If mbIgnore Then Exit Sub

    GlobalParms.Save "MaxConsecutiveSound", "NumVal", CInt(Me!Combo32)
    GlobalParms.Save "SPAMOptionsChanged", "TrueFalse", True
    ShowMessage "Change saved", 500, Me

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub Combo33_Click()

On Error GoTo ErrorTrap

    If mbIgnore Then Exit Sub

    GlobalParms.Save "MaxConsecutivePlatform", "NumVal", CInt(Me!Combo33)
    GlobalParms.Save "SPAMOptionsChanged", "TrueFalse", True
    ShowMessage "Change saved", 500, Me

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

