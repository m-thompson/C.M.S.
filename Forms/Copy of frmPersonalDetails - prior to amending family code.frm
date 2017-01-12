VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPersonalDetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.M.S. Personal Details"
   ClientHeight    =   6870
   ClientLeft      =   420
   ClientTop       =   990
   ClientWidth     =   8235
   Icon            =   "frmPersonalDetails.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTelNoSearch 
      Caption         =   "&Tel Search"
      Height          =   300
      Left            =   4755
      TabIndex        =   81
      Top             =   1515
      Width           =   930
   End
   Begin VB.CommandButton cmdContact 
      Caption         =   "Contact"
      Height          =   300
      Left            =   4755
      TabIndex        =   77
      Top             =   1845
      Width           =   930
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   314
      Left            =   5300
      TabIndex        =   46
      ToolTipText     =   "Find Next"
      Top             =   315
      Width           =   150
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   314
      Left            =   5160
      TabIndex        =   45
      ToolTipText     =   "Find Previous"
      Top             =   315
      Width           =   150
   End
   Begin TabDlg.SSTab tabPeopleTabs 
      Height          =   4395
      Left            =   225
      TabIndex        =   51
      Top             =   2280
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   7752
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Personal Details"
      TabPicture(0)   =   "frmPersonalDetails.frx":0442
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "chkVisitingSpeaker"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmbGender"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "chkAnointed"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtFirstName"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtMiddleName"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtLastName"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtDOB"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "chkActive"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmbInfirmity"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblVisiting"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label14"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label34"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label32"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label30"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label28"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label24"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label22"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Family Details"
      TabPicture(1)   =   "frmPersonalDetails.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstParents"
      Tab(1).Control(1)=   "cmbAddFromList"
      Tab(1).Control(2)=   "lstChildren"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "txtSpouse"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "optgrpEditWhat"
      Tab(1).Control(5)=   "Label75"
      Tab(1).Control(6)=   "lblAddFamily"
      Tab(1).Control(7)=   "Label59"
      Tab(1).Control(8)=   "Label57"
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Congregation Roles"
      TabPicture(2)   =   "frmPersonalDetails.frx":047A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "cmdSuspendDates"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "cmdTMSStudents"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdSPAM"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdBookGroups"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdFieldService"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmdElderServants"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdPublicMeetingPersonnel"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cmdVisitingSpeakers"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmdServiceMtgs"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "frmMinistry"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).ControlCount=   10
      TabCaption(3)   =   "Address Details"
      TabPicture(3)   =   "frmPersonalDetails.frx":0496
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame3"
      Tab(3).Control(1)=   "cmbNames"
      Tab(3).Control(2)=   "Label13"
      Tab(3).ControlCount=   3
      Begin VB.CommandButton frmMinistry 
         Caption         =   "&Field Ministry..."
         Height          =   450
         Left            =   2940
         TabIndex        =   17
         Top             =   1155
         Width           =   1845
      End
      Begin VB.CheckBox chkVisitingSpeaker 
         Alignment       =   1  'Right Justify
         Height          =   210
         Left            =   -71460
         TabIndex        =   78
         Top             =   2490
         Width           =   240
      End
      Begin VB.ComboBox cmbGender 
         Height          =   315
         Left            =   -73335
         TabIndex        =   5
         Text            =   "cmbGender"
         Top             =   2010
         Width           =   975
      End
      Begin VB.CommandButton cmdServiceMtgs 
         Caption         =   "Ser&vice Meetings..."
         Height          =   450
         Left            =   1095
         TabIndex        =   24
         Top             =   2955
         Width           =   1845
      End
      Begin VB.CommandButton cmdVisitingSpeakers 
         Caption         =   "&Visiting Speakers..."
         Height          =   450
         Left            =   2940
         TabIndex        =   23
         Top             =   2505
         Width           =   1845
      End
      Begin VB.CommandButton cmdPublicMeetingPersonnel 
         Caption         =   "&Public Meetings..."
         Height          =   450
         Left            =   1095
         TabIndex        =   22
         Top             =   2505
         Width           =   1845
      End
      Begin VB.CheckBox chkAnointed 
         Caption         =   "Active?:"
         Height          =   240
         Left            =   -73335
         TabIndex        =   9
         Top             =   3780
         Width           =   260
      End
      Begin VB.Frame Frame3 
         Height          =   3375
         Left            =   -74850
         TabIndex        =   67
         Top             =   855
         Width           =   5655
         Begin VB.CommandButton cmdSearchForTelNo 
            Caption         =   "Search"
            Height          =   300
            Left            =   3615
            TabIndex        =   80
            Top             =   2595
            Width           =   795
         End
         Begin VB.TextBox txtAddress1 
            Height          =   314
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   27
            Top             =   240
            Width           =   3729
         End
         Begin VB.TextBox txtAddress2 
            Height          =   314
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   28
            Top             =   630
            Width           =   3729
         End
         Begin VB.TextBox txtAddress3 
            Height          =   314
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   29
            Top             =   1020
            Width           =   3729
         End
         Begin VB.TextBox txtAddress4 
            Height          =   314
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   30
            Top             =   1425
            Width           =   3729
         End
         Begin VB.TextBox txtPostcode 
            Height          =   314
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   31
            Top             =   1815
            Width           =   1020
         End
         Begin VB.TextBox txtHomePhone 
            Height          =   314
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   32
            Top             =   2190
            Width           =   1920
         End
         Begin VB.TextBox txtMobile 
            Height          =   314
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   33
            Top             =   2580
            Width           =   1920
         End
         Begin VB.TextBox txtEmail 
            Height          =   314
            Left            =   1665
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   34
            Top             =   2970
            Width           =   3729
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Address 1:"
            Height          =   240
            Left            =   240
            TabIndex        =   75
            Top             =   270
            Width           =   1365
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Address 2:"
            Height          =   240
            Left            =   240
            TabIndex        =   74
            Top             =   660
            Width           =   1365
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Address 3:"
            Height          =   240
            Left            =   240
            TabIndex        =   73
            Top             =   1050
            Width           =   1365
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Address 4:"
            Height          =   240
            Left            =   240
            TabIndex        =   72
            Top             =   1455
            Width           =   1365
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Postcode:"
            Height          =   240
            Left            =   240
            TabIndex        =   71
            Top             =   1845
            Width           =   1365
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Home Phone:"
            Height          =   240
            Left            =   240
            TabIndex        =   70
            Top             =   2220
            Width           =   1365
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Mobile Phone:"
            Height          =   240
            Left            =   240
            TabIndex        =   69
            Top             =   2610
            Width           =   1365
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Email Address:"
            Height          =   240
            Left            =   240
            TabIndex        =   68
            Top             =   3000
            Width           =   1365
         End
      End
      Begin VB.ComboBox cmbNames 
         Height          =   315
         Left            =   -73185
         TabIndex        =   26
         Text            =   "cmbNames"
         Top             =   495
         Width           =   3750
      End
      Begin VB.TextBox txtFirstName 
         Height          =   314
         Left            =   -73335
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   2
         Top             =   705
         Width           =   3729
      End
      Begin VB.TextBox txtMiddleName 
         Height          =   314
         Left            =   -73335
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1140
         Width           =   3729
      End
      Begin VB.TextBox txtLastName 
         Height          =   314
         Left            =   -73335
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1575
         Width           =   3729
      End
      Begin VB.TextBox txtDOB 
         Alignment       =   2  'Center
         Height          =   314
         Left            =   -73335
         Locked          =   -1  'True
         MaxLength       =   10
         TabIndex        =   7
         Top             =   2865
         Width           =   954
      End
      Begin VB.CheckBox chkActive 
         Caption         =   "Anointed?:"
         Height          =   240
         Left            =   -73335
         TabIndex        =   6
         Top             =   2475
         Width           =   260
      End
      Begin VB.ComboBox cmbInfirmity 
         Height          =   315
         Left            =   -73335
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3300
         Width           =   2025
      End
      Begin VB.CommandButton cmdElderServants 
         Caption         =   "E&lders && Servants..."
         Height          =   450
         Left            =   2940
         TabIndex        =   19
         Top             =   1605
         Width           =   1845
      End
      Begin VB.CommandButton cmdFieldService 
         Caption         =   "&Publisher Status..."
         Height          =   450
         Left            =   1095
         TabIndex        =   16
         Top             =   1155
         Width           =   1845
      End
      Begin VB.CommandButton cmdBookGroups 
         Caption         =   "Boo&k Group..."
         Height          =   450
         Left            =   2940
         TabIndex        =   25
         Top             =   2955
         Width           =   1845
      End
      Begin VB.CommandButton cmdSPAM 
         Caption         =   "A&ttendants..."
         Height          =   450
         Left            =   1095
         TabIndex        =   20
         Top             =   2055
         Width           =   1845
      End
      Begin VB.CommandButton cmdTMSStudents 
         Caption         =   "&School..."
         Height          =   450
         Left            =   1095
         TabIndex        =   18
         Top             =   1605
         Width           =   1845
      End
      Begin VB.CommandButton cmdSuspendDates 
         Caption         =   "&Suspend ..."
         Height          =   450
         Left            =   2940
         TabIndex        =   21
         Top             =   2055
         Width           =   1845
      End
      Begin VB.ListBox lstParents 
         Height          =   450
         Left            =   -73860
         TabIndex        =   15
         Top             =   3315
         Width           =   3974
      End
      Begin VB.ComboBox cmbAddFromList 
         Height          =   315
         Left            =   -73860
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2385
         Visible         =   0   'False
         Width           =   3974
      End
      Begin VB.ListBox lstChildren 
         Height          =   1425
         Left            =   -73860
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   900
         Width           =   3974
      End
      Begin VB.TextBox txtSpouse 
         Height          =   310
         Left            =   -73860
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   540
         Width           =   3974
      End
      Begin VB.Frame optgrpEditWhat 
         Height          =   405
         Left            =   -73785
         TabIndex        =   59
         Top             =   2655
         Visible         =   0   'False
         Width           =   2385
         Begin VB.OptionButton optEditChild 
            Caption         =   "Child"
            Height          =   240
            Left            =   1425
            TabIndex        =   14
            Top             =   135
            Width           =   260
         End
         Begin VB.OptionButton optEditSpouse 
            Caption         =   "Spouse"
            Height          =   240
            Left            =   165
            TabIndex        =   13
            Top             =   135
            Width           =   260
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Child"
            Height          =   240
            Left            =   1845
            TabIndex        =   65
            Top             =   135
            Width           =   615
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Spouse"
            Height          =   240
            Left            =   585
            TabIndex        =   64
            Top             =   135
            Width           =   615
         End
      End
      Begin VB.Label lblVisiting 
         Caption         =   "Visiting Speaker?"
         Height          =   210
         Left            =   -72750
         TabIndex        =   79
         Top             =   2490
         Width           =   1305
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Anointed?"
         Height          =   240
         Left            =   -74775
         TabIndex        =   76
         Top             =   3765
         Width           =   1530
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Link Address To:"
         Height          =   240
         Left            =   -74595
         TabIndex        =   66
         Top             =   555
         Width           =   1365
      End
      Begin VB.Label Label75 
         BackStyle       =   0  'Transparent
         Caption         =   "Parents:"
         Height          =   240
         Left            =   -74850
         TabIndex        =   63
         Top             =   3345
         Width           =   660
      End
      Begin VB.Label lblAddFamily 
         BackStyle       =   0  'Transparent
         Caption         =   "Add from list:"
         Height          =   240
         Left            =   -74850
         TabIndex        =   62
         Top             =   2445
         Width           =   975
      End
      Begin VB.Label Label59 
         BackStyle       =   0  'Transparent
         Caption         =   "Children"
         Height          =   240
         Left            =   -74850
         TabIndex        =   61
         Top             =   945
         Width           =   645
      End
      Begin VB.Label Label57 
         BackStyle       =   0  'Transparent
         Caption         =   "Spouse"
         Height          =   240
         Left            =   -74850
         TabIndex        =   60
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "In Local Cong?"
         Height          =   240
         Left            =   -74775
         TabIndex        =   58
         Top             =   2490
         Width           =   1530
      End
      Begin VB.Label Label34 
         BackStyle       =   0  'Transparent
         Caption         =   "Infirmity Level:"
         Height          =   240
         Left            =   -74775
         TabIndex        =   57
         Top             =   3360
         Width           =   1530
      End
      Begin VB.Label Label32 
         BackStyle       =   0  'Transparent
         Caption         =   "DOB:"
         Height          =   240
         Left            =   -74775
         TabIndex        =   56
         Top             =   2925
         Width           =   1470
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         Height          =   240
         Left            =   -74775
         TabIndex        =   55
         Top             =   2010
         Width           =   1425
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name:"
         Height          =   240
         Left            =   -74760
         TabIndex        =   54
         Top             =   1605
         Width           =   1470
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Name:"
         Height          =   240
         Left            =   -74760
         TabIndex        =   53
         Top             =   1185
         Width           =   1425
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "First Name:"
         Height          =   240
         Left            =   -74760
         TabIndex        =   52
         Top             =   735
         Width           =   1365
      End
   End
   Begin VB.ListBox lstNames 
      Height          =   1425
      Left            =   1125
      TabIndex        =   1
      Top             =   720
      Width           =   3572
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   567
      Left            =   6585
      TabIndex        =   36
      Top             =   1005
      Width           =   1304
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Enabled         =   0   'False
      Height          =   567
      Left            =   6600
      TabIndex        =   42
      Top             =   5925
      Width           =   1304
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   567
      Left            =   6585
      TabIndex        =   37
      Top             =   1695
      Width           =   1304
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   567
      Left            =   6585
      TabIndex        =   35
      Top             =   330
      Width           =   1304
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   567
      Left            =   6600
      TabIndex        =   41
      Top             =   4860
      Width           =   1304
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New"
      Height          =   567
      Left            =   6600
      TabIndex        =   39
      Top             =   3405
      Width           =   1304
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Edit"
      Height          =   567
      Left            =   6600
      TabIndex        =   40
      Top             =   4125
      Width           =   1304
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Enabled         =   0   'False
      Height          =   567
      Left            =   6600
      TabIndex        =   38
      Top             =   2700
      Width           =   1304
   End
   Begin VB.CheckBox chkActiveOnly 
      Caption         =   "&Cong Only"
      Height          =   240
      Left            =   4785
      TabIndex        =   43
      Top             =   870
      Width           =   1080
   End
   Begin VB.TextBox txtSearch 
      Height          =   314
      Left            =   1125
      TabIndex        =   0
      Top             =   315
      Width           =   3572
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "&Go"
      Enabled         =   0   'False
      Height          =   314
      Left            =   4755
      TabIndex        =   44
      ToolTipText     =   "Search"
      Top             =   315
      Width           =   420
   End
   Begin VB.Frame Frame1 
      Height          =   2310
      Left            =   6465
      TabIndex        =   49
      Top             =   105
      Width           =   1560
   End
   Begin VB.Frame Frame2 
      Height          =   4200
      Left            =   6465
      TabIndex        =   50
      Top             =   2505
      Width           =   1560
   End
   Begin VB.Label Name_Label 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   240
      Left            =   510
      TabIndex        =   47
      Top             =   765
      Width           =   540
   End
   Begin VB.Label Label53 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Find Name"
      Height          =   240
      Left            =   210
      TabIndex        =   48
      Top             =   345
      Width           =   840
   End
   Begin VB.Menu mnuCongregations 
      Caption         =   "Congregations"
      Begin VB.Menu mnuAddCong 
         Caption         =   "Add Congregation"
      End
      Begin VB.Menu mnuEditCong 
         Caption         =   "Edit Congregation"
      End
   End
End
Attribute VB_Name = "frmPersonalDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rstNameAddress As Recordset, rstNameAddress2 As Recordset
Dim rstMarriage As Recordset, rstChildren As Recordset, rstResponsibilities As Recordset
Dim NewRec As Boolean, UpdateRec As Boolean, HoldID As Integer
Dim NewSpouse As Boolean, ReplaceSpouse As Boolean
Dim NewChild As Boolean, TheNewSpouse As Integer, Initialising As Boolean
Dim StoredFirstName As String, StoredMiddleName As String, StoredLastName As String
Dim TheSelectedPerson As Long, PersonIsActive As Boolean, SearchActive As Boolean
Dim rstSearch As Recordset, DoNotTriggerAddressStuff As Boolean, DoNotTrigger As Boolean


Private Sub chkActiveOnly_Click()

On Error GoTo ErrorTrap

    'include inactive Name & Address Records in the lstNames? This is decided by chkActiveOnly.
      
    If chkActiveOnly.value = vbChecked Then
                    
        PopulateNamesList ShowActiveOnly:=True
        
        Set rstNameAddress = CMSDB.OpenRecordset("SELECT * " & _
                                        "FROM tblNameAddress " & _
                                        "WHERE Active = TRUE " & _
                                        "ORDER BY LastName, FirstName" _
                                        , dbOpenDynaset)
    Else

        PopulateNamesList ShowActiveOnly:=False
        
        Set rstNameAddress = CMSDB.OpenRecordset("SELECT * " & _
                                "FROM tblNameAddress " & _
                                "ORDER BY LastName, FirstName" _
                                , dbOpenDynaset)

    End If
    
    TextFieldGotFocus txtSearch, True

    Select Case tabPeopleTabs.Tab
    Case 0: Call SetUpPersonalTab
    Case 1: Call SetUpFamilyTab
    Case 2: Call SetUpResponsibilitiesTab
    End Select

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub




Private Sub chkVisitingSpeaker_Click()

On Error GoTo ErrorTrap

    Call ChangedFieldsPersonal
    
    If chkVisitingSpeaker.value = vbChecked Then
        DoNotTrigger = True
        chkActive.value = vbUnchecked
        DoNotTrigger = False
    End If

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmbAddFromList_Click()
    
    If cmbAddFromList.ListIndex > -1 Then
        cmdApply.Enabled = True
    Else
        cmdApply.Enabled = False
    End If
        
End Sub

Private Sub cmbGender_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap

    AutoCompleteComboKeyDown cmbGender, KeyCode

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmbGender_KeyPress(KeyAscii As Integer)

On Error GoTo ErrorTrap

    AutoCompleteCombo cmbGender, KeyAscii

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmbInfirmity_Click()
    ChangedFieldsPersonal
End Sub


Private Sub cmbNames_Click()

    If DoNotTriggerAddressStuff Then Exit Sub
    
    If UpdateRec Then
        cmdApply.Enabled = True
    End If
    
    PopulateFieldsAddress
    
End Sub

Private Sub cmbNames_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 46 Then
        cmbNames.ListIndex = -1
    End If
    
End Sub

Private Sub cmbNames_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorTrap

    AutoCompleteCombo Me!cmbNames, KeyAscii

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdApply_Click()

On Error GoTo ErrorTrap

    Dim SaveAndExit As Boolean
    
    Select Case tabPeopleTabs.Tab
    Case 0: Call ApplyChangesPersonal(SaveAndExit)
    Case 1: Call ApplyChangesFamily(SaveAndExit)
    Case 3: Call ApplyChangesAddress(SaveAndExit)
    End Select


    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub cmdBookGroups_Click()
On Error GoTo ErrorTrap

    lnkPersonID = Me!lstNames.ItemData(Me!lstNames.ListIndex)
    frmBookGroupMembers.Show vbModal

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdBrowse_Click()

On Error GoTo ErrorTrap


    Select Case tabPeopleTabs.Tab
    Case 0: Call GoToBrowseModePersonal
    Case 1: Call GoToBrowseModeFamily
    Case 3: Call GoToBrowseModeAddress
    End Select
    

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub cmdCancel_Click()

On Error GoTo ErrorTrap

    Unload Me

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub cmdContact_Click()

On Error GoTo ErrorTrap

    With frmContactDetails
    .PersonID = TheSelectedPerson
    .MessageBody = ""
    .MessageTitle = ""
    .Show vbModal, Me
    End With
    
    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdDelete_Click()

On Error GoTo ErrorTrap

    Select Case tabPeopleTabs.Tab
    Case 0: Call DeleteRecPersonal
    Case 1: Call DeleteRecFamily
    End Select

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub



Private Sub cmdElderServants_Click()
On Error GoTo ErrorTrap

    frmEldersAndServants.FormPersonID = TheSelectedPerson
    frmEldersAndServants.Show vbModal

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub cmdFieldService_Click()
On Error GoTo ErrorTrap

    frmFieldServiceRoles.FormPersonID = TheSelectedPerson
    frmFieldServiceRoles.Show vbModal

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub cmdGo_Click()

On Error GoTo ErrorTrap

    GoSearchForName

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub GoSearchForName()
Dim TheSearchString As String

On Error GoTo ErrorTrap


    If Len(txtSearch) > 0 Or Not IsNull(txtSearch) Then
        TheSearchString = DoubleUpSingleQuotes(txtSearch)
        If chkActiveOnly.value = vbChecked Then
            Set rstSearch = CMSDB.OpenRecordset("SELECT ID " & _
                                    "FROM tblNameAddress " & _
                                    "WHERE InStr(1, LastName, " & "'" & TheSearchString & "'" & ", 1) > 0" & _
                                    "OR    InStr(1, FirstName, " & "'" & TheSearchString & "'" & ", 1) > 0" & _
                                    " AND Active = TRUE" & _
                                    " ORDER BY LastName, FirstName" _
                                    , dbOpenDynaset)
        Else
            Set rstSearch = CMSDB.OpenRecordset("SELECT ID " & _
                            "FROM tblNameAddress " & _
                            "WHERE InStr(1, LastName, " & "'" & TheSearchString & "'" & ", 1) > 0" & _
                            "OR    InStr(1, FirstName, " & "'" & TheSearchString & "'" & ", 1) > 0" & _
                            " ORDER BY LastName, FirstName" _
                            , dbOpenDynaset)
        End If


        If rstSearch.BOF Then
            MsgBox "Name not found.", vbOKOnly + vbExclamation, AppName
            txtSearch.SetFocus
            SearchActive = False
            cmdNext.Enabled = False
            cmdPrev.Enabled = False
        Else
            SearchActive = True
            cmdNext.Enabled = True
            cmdPrev.Enabled = True
            HoldID = rstSearch!ID
            lstNames.SetFocus
            MoveToNextRecPersonal
            PopulateFieldsFamily
        End If

    End If


    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub cmdNew_Click()

On Error GoTo ErrorTrap

    
    Select Case tabPeopleTabs.Tab
    Case 0: Call GoToNewModePersonal
    End Select


    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub cmdNext_Click()
    SearchNext
End Sub

Private Sub cmdOK_Click()

On Error GoTo ErrorTrap

    Dim SaveAndExit As Boolean
    
    If NewRec Or UpdateRec Then
    
        Select Case tabPeopleTabs.Tab
        Case 0: Call ApplyChangesPersonal(SaveAndExit)
        Case 1: Call ApplyChangesFamily(SaveAndExit)
        Case 3: Call ApplyChangesAddress(SaveAndExit)
        End Select
        
        If SaveAndExit Then
            Unload Me
        End If
        
    Else
        Unload Me
    End If
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub cmdPrev_Click()
    SearchPrev
End Sub

Private Sub cmdPublicMeetingPersonnel_Click()

On Error GoTo ErrorTrap

    frmPublicMtgPersonnel.FormPersonID = TheSelectedPerson
    frmPublicMtgPersonnel.Show vbModal, Me

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub cmdRefresh_Click()
'Repopulate text fields based on selection in lstNames

On Error GoTo ErrorTrap


    Select Case tabPeopleTabs.Tab
    Case 0: RefreshPersonal
    End Select

    Exit Sub
ErrorTrap:
    EndProgram
    
    
End Sub

Private Sub cmdSearchForTelNo_Click()
On Error GoTo ErrorTrap

    frmSearchTelNo.Show vbModal, Me


    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdServiceMtgs_Click()

On Error GoTo ErrorTrap

    frmServiceMtgPersonnel.FormPersonID = TheSelectedPerson
    frmServiceMtgPersonnel.Show vbModal, Me

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdSPAM_Click()
    frmAttendantsDetails.FormPersonID = TheSelectedPerson
    frmAttendantsDetails.Show vbModal
    GoToBrowseModeRoles
End Sub

Private Sub cmdSuspendDates_Click()

On Error GoTo ErrorTrap

    lnkCongNo = CInt(GlobalDefaultCong)
    If Me!lstNames.ListIndex > -1 Then
        lnkPersonID = Me!lstNames.ItemData(Me!lstNames.ListIndex)
    Else
        lnkPersonID = -1
    End If
    lnkLimitTaskCatsSQL = "" 'allow user to see all task categories
    lnkLimitTaskSubCatsSQL = "" 'allow user to see all task subcategories
    lnkLimitTaskCatsSQL2 = ""
    lnkLimitTaskSubCatsSQL2 = ""
    frmSuspendDatesMaint.Show vbModal

    Exit Sub
ErrorTrap:
    EndProgram
    
        
End Sub

Private Sub cmdTelNoSearch_Click()

On Error GoTo ErrorTrap

    frmSearchTelNo.Show vbModal, Me

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Private Sub cmdTMSStudents_Click()

Dim StorePerson As Integer

    frmTMSStudentDetails.FormPersonID = TheSelectedPerson
    frmTMSStudentDetails.Show vbModal
    
    SetUpNameAddressRecSets 'Recsets may have been lost if report run
    
    StorePerson = lstNames.ListIndex
    
    If StorePerson > -1 Then
        lstNames.ListIndex = StorePerson
        tabPeopleTabs.Tab = 2
    End If
    
End Sub

Private Sub cmdUpdate_Click()

On Error GoTo ErrorTrap

    If rstNameAddress.BOF Then Exit Sub
    
    Select Case tabPeopleTabs.Tab
    Case 0: Call GoToUpdateModePersonal
    Case 1: Call GoToUpdateModeFamily
    Case 3: Call GoToUpdateModeAddress
    End Select

    Exit Sub
ErrorTrap:
    EndProgram

    
End Sub



Private Sub cmdVisitingSpeakers_Click()
On Error GoTo ErrorTrap

    frmVisitingSpeakers.FormPersonID = TheSelectedPerson
    frmVisitingSpeakers.Show vbModal, Me

    Exit Sub
ErrorTrap:
    Call EndProgram

End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorTrap

    If KeyCode = vbKeyS And ((Shift And vbCtrlMask) > 0) Then
        ShowAllOpenFormsExceptMenus
    End If
    
    If KeyCode = vbKeyH And ((Shift And vbCtrlMask) > 0) Then
        HideAllOpenFormsExceptMenus
    End If
    
    
    Exit Sub
    
ErrorTrap:
    EndProgram
End Sub


Private Sub Form_Load()

On Error GoTo ErrorTrap

    tabPeopleTabs.Tab = 0
    
    PopulateNamesList ShowActiveOnly:=True
            
    SetUpNameAddressRecSets
    
    FillInfirmityCombo
    
    cmbGender.AddItem "Female"
    cmbGender.AddItem "Male"
                                                                                            
    If Not rstNameAddress.BOF Then
        SetUpPersonalTab
        
        SetUpFamilyTab
        
        SetUpResponsibilitiesTab
        
        chkActiveOnly.value = vbChecked
    Else
        tabPeopleTabs.TabEnabled(1) = False
        tabPeopleTabs.TabEnabled(2) = False
    End If
    
    EnforceSecurity
    
    lnkPersDtlsFormIsOpen = True 'Done for the sake of frmCongStructure
    
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PopulateNamesList(ShowActiveOnly As Boolean)

On Error GoTo ErrorTrap

    HandleListBox.PopulateListBox frmPersonalDetails!lstNames, _
        "SELECT DISTINCTROW tblNameAddress.ID, " & _
        "                   tblNameAddress.FirstName & ' ' & " & _
        "                   tblNameAddress.MiddleName, " & _
        "                   tblNameAddress.LastName " & _
        "FROM tblNameAddress " & _
        IIf(ShowActiveOnly, "WHERE Active = TRUE ", "") & _
        " ORDER BY tblNameAddress.LastName, tblNameAddress.FirstName" _
        , CMSDB, 0, ", ", True, 2, 1

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    lnkPersDtlsFormIsOpen = False
    CongregationMember.Class_Initialize
    BringForwardMainMenuWhenItsTheLastFormOpen
    
    If SearchActive Then
        rstSearch.Close
    End If
    
    Cancel = False
End Sub

Private Sub frmMinistry_Click()

On Error GoTo ErrorTrap

    frmCongStats.ThePerson = FormPersonID
    frmCongStats.Show vbModeless, Me


    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Sub lstChildren_GotFocus()

On Error GoTo ErrorTrap

    ListChildren
    optEditChild.value = True

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub lstNames_Click()

On Error GoTo ErrorTrap

    
   ' Find the record that matches lstNames.

    SelectFromLstNamesPersonal
    SelectFromLstNamesFamily
    SelectFromLstNamesAddress
       
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub FillInfirmityCombo()

On Error GoTo ErrorTrap

    With cmbInfirmity
    
    .AddItem "No health problems"
    .AddItem "Minor health problems"
    .AddItem "Major health problems"
    .AddItem "Severe disability"
    
    End With
       
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub PopulateFieldsPersonal()

On Error GoTo ErrorTrap

    If rstNameAddress.BOF Then Exit Sub
   
    txtFirstName = rstNameAddress!FirstName
    txtLastName = rstNameAddress!LastName
    If Not IsNull(rstNameAddress!MiddleName) Then
        txtMiddleName = rstNameAddress!MiddleName
    Else
        txtMiddleName = ""
    End If
    cmbGender.ListIndex = IIf(rstNameAddress!GenderMF = "M", 1, 0)
    txtFirstName = rstNameAddress!FirstName
    txtDOB = IIf(rstNameAddress!DOB > 0, rstNameAddress!DOB, "")
'    txtInfirmityLevel = rstNameAddress![InfirmityLevel0-6]
    If rstNameAddress![InfirmityLevel0-6] > cmbInfirmity.ListCount - 1 Then
        cmbInfirmity.ListIndex = cmbInfirmity.ListCount - 1
    Else
        cmbInfirmity.ListIndex = rstNameAddress![InfirmityLevel0-6]
    End If
    
    If CongregationMember.IsVisitingSpeaker(rstNameAddress!ID) Then
        chkVisitingSpeaker.value = vbChecked
    Else
        chkVisitingSpeaker.value = vbUnchecked
    End If
    
    '
    'Store the name fields for use in checking whether user has changed them
    '
    StoredFirstName = txtFirstName.text
    StoredMiddleName = txtMiddleName.text
    StoredLastName = txtLastName.text
    
    Select Case rstNameAddress!Active
    Case True:
        chkActive = vbChecked
        PersonIsActive = True
    Case False:
        chkActive = vbUnchecked
        PersonIsActive = False
    End Select

    Select Case rstNameAddress!Anointed
    Case True:
        chkAnointed = vbChecked
    Case False:
        chkAnointed = vbUnchecked
    End Select
        
    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub lstNames_DblClick()

On Error GoTo ErrorTrap


    Dim CancelOption As Integer

    If MsgBox("Do you want to edit this record? ", vbYesNo + vbQuestion + vbDefaultButton2, AppName) = vbYes Then
        GoToBrowseModeFamily
'        GoToBrowseModeRoles
        tabPeopleTabs.Tab = 0
        GoToUpdateModePersonal
    End If

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub lstNames_KeyPress(KeyAscii As Integer)
On Error GoTo ErrorTrap
    
    If KeyAscii = 13 Then
        SearchNext
    End If
    
    Exit Sub
    
ErrorTrap:
    EndProgram


End Sub

Private Sub mnuAddCong_Click()
On Error GoTo ErrorTrap
    
    frmEditCongList.AddMode = True
    frmEditCongList.Show vbModal
    
    Exit Sub
    
ErrorTrap:
    EndProgram

End Sub
Private Sub SearchNext()
On Error GoTo ErrorTrap
    
    If SearchActive Then
        With rstSearch
        .MoveNext
        If .EOF Then
            .MoveFirst
        End If
        HoldID = rstSearch!ID
        lstNames.SetFocus
        MoveToNextRecPersonal
        PopulateFieldsFamily
        End With
    End If
    
    Exit Sub
    
ErrorTrap:
    EndProgram

End Sub
Private Sub SearchPrev()
On Error GoTo ErrorTrap
    
    If SearchActive Then
        With rstSearch
        .MovePrevious
        If .BOF Then
            .MoveLast
        End If
        HoldID = rstSearch!ID
        lstNames.SetFocus
        MoveToNextRecPersonal
        PopulateFieldsFamily
        End With
    End If
    
    Exit Sub
    
ErrorTrap:
    EndProgram

End Sub

Private Sub mnuEditCong_Click()
On Error GoTo ErrorTrap
    
    frmEditCongList.AddMode = False
    frmEditCongList.SelectedCongForEdit = CInt(GlobalDefaultCong)
    frmEditCongList.Show vbModal
    
    Exit Sub
    
ErrorTrap:
    EndProgram


End Sub

Private Sub optEditChild_Click()

On Error GoTo ErrorTrap

    If optEditChild.value = True Then
        ListChildren
    End If

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub optEditSpouse_Click()

On Error GoTo ErrorTrap


    If optEditSpouse.value = True Then
        ListSpouses
    End If

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub ListChildren()
Dim ChildrenSQL As String


On Error GoTo ErrorTrap

'List all possible children that don't already have a parent of the gender in txtGender. Also
' age of children listed should be less than that of parent in lstNames
    
    lblAddFamily.Caption = "Add Child:"
    
   
    ChildrenSQL = "SELECT tblNameAddress.ID, " & _
                         "tblNameAddress.FirstName & ' ' & " & _
                         "tblNameAddress.MiddleName & ' ' & " & _
                         "tblNameAddress.LastName " & _
                  "FROM tblNameAddress " & _
                  "WHERE ID NOT IN (SELECT tblChildren.Child " & _
                                    "FROM tblChildren " & _
                                    "WHERE tblChildren.Parent IN (SELECT tblNameAddress.ID " & _
                                                                 "FROM tblNameAddress " & _
                                                                 "WHERE GenderMF = '" & _
                                                                 IIf(cmbGender.ListIndex = 1, "M", "F") & "')) " & _
                  "AND DOB > (SELECT DOB " & _
                             "FROM tblNameAddress " & _
                             "WHERE ID = " & lstNames.ItemData(lstNames.ListIndex) & ") " & _
                  "ORDER BY LastName, FirstName"

    
    HandleListBox.PopulateListBox frmPersonalDetails!cmbAddFromList, ChildrenSQL, CMSDB, 0, "", True, 1
    
'Prepare 'update mode' flag
    NewChild = True
    NewSpouse = False
    ReplaceSpouse = False


    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub ListSpouses()
Dim SpouseSQL As String, Gender As String


On Error GoTo ErrorTrap

    
'List all possible spouses
    
    lblAddFamily.Caption = "Add Spouse:"
    
    If cmbGender.ListIndex = 0 Then
        Gender = "M"
    Else
        Gender = "F"
    End If
     
    SpouseSQL = "SELECT tblNameAddress.ID, " & _
                         "tblNameAddress.FirstName & ' ' &" & _
                         "tblNameAddress.MiddleName &  ' ' &" & _
                         "tblNameAddress.LastName " & _
                "FROM tblNameAddress " & _
                "WHERE tblNameAddress.GenderMF = '" & Gender & "'" & _
               " AND ID NOT IN (SELECT tblMarriage.Spouse " & _
                           "FROM tblMarriage) " & _
                "ORDER BY LastName, FirstName"
        
    HandleListBox.PopulateListBox frmPersonalDetails!cmbAddFromList, SpouseSQL, CMSDB, 0, "", True, 1
    
' Prepare the 'type of update' flags. Decide whether there is an existing Spouse, in which case this should be replaced
'   if an update is made

    If IsNull(txtSpouse) Or Len(txtSpouse) = 0 Then
        NewSpouse = True
        ReplaceSpouse = False
    Else
        NewSpouse = False
        ReplaceSpouse = True
    End If
    
    NewChild = False
        

    Exit Sub
ErrorTrap:
    EndProgram
End Sub




Private Sub tabPeopleTabs_Click(PreviousTab As Integer)

On Error GoTo ErrorTrap

    If rstNameAddress.BOF Then Exit Sub
    
    GoToBrowseModePersonal
    GoToBrowseModeFamily
    GoToBrowseModeAddress

    Select Case tabPeopleTabs.Tab
    Case 0: cmdNew.Enabled = True
    Case 1: cmdNew.Enabled = False
    Case 2: GoToBrowseModeRoles
    Case 3: cmdNew.Enabled = False
    End Select

    Select Case tabPeopleTabs.Tab
    Case 0: cmdUpdate.Enabled = True
    Case 1: cmdUpdate.Enabled = True
    Case 3: cmdUpdate.Enabled = True
    End Select

    Select Case tabPeopleTabs.Tab
    Case 0: cmdDelete.Enabled = True
    Case 1: cmdDelete.Enabled = True
    End Select

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub txtAddress1_Change()
    If UpdateRec Then
        cmdApply.Enabled = True
    End If
End Sub

Private Sub txtAddress1_GotFocus()
    TextFieldGotFocus txtAddress1
End Sub

Private Sub txtAddress2_Change()
    If UpdateRec Then
        cmdApply.Enabled = True
    End If
End Sub

Private Sub txtAddress2_GotFocus()
    TextFieldGotFocus txtAddress2
End Sub

Private Sub txtAddress3_Change()
    If UpdateRec Then
        cmdApply.Enabled = True
    End If

End Sub

Private Sub txtAddress3_GotFocus()
    TextFieldGotFocus txtAddress3
End Sub

Private Sub txtAddress4_Change()
    If UpdateRec Then
        cmdApply.Enabled = True
    End If

End Sub

Private Sub txtAddress4_GotFocus()
    TextFieldGotFocus txtAddress4
End Sub

Private Sub txtDOB_GotFocus()
    TextFieldGotFocus txtDOB
End Sub

Private Sub txtDOB_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorTrap

'Must be numeric. Allow Backspace (8) and forward-slash (47). Delete and arrow keys seem to be allowed by default.

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 47 Then
        KeyAscii = 0
    End If
    Exit Sub
ErrorTrap:
    EndProgram


End Sub

Private Sub txtEmail_Change()
    If UpdateRec Then
        cmdApply.Enabled = True
    End If

End Sub

Private Sub txtEmail_GotFocus()
    TextFieldGotFocus txtEmail
End Sub

Private Sub txtFirstName_GotFocus()
    TextFieldGotFocus txtFirstName
End Sub

Private Sub txtFirstName_LostFocus()

On Error GoTo ErrorTrap


    txtFirstName = Trim(txtFirstName)
    'txtFirstName = DelSubstr(txtFirstName, "'", False)

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub




'Private Sub txtInfirmityLevel_LostFocus()
'
'On Error GoTo ErrorTrap
'
'
'    txtInfirmityLevel = Trim(txtInfirmityLevel)
'
'    Exit Sub
'ErrorTrap:
'    EndProgram
'
'
'End Sub

Private Sub txtDOB_LostFocus()
 
On Error GoTo ErrorTrap

   txtDOB = Trim(txtDOB)

    If NewRec Or UpdateRec Then
       If IsDate(txtDOB) Then
           txtDOB = Format(txtDOB, "Short Date")
       End If
    End If
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub txtHomePhone_Change()
    If UpdateRec Then
        cmdApply.Enabled = True
    End If

End Sub

Private Sub txtHomePhone_GotFocus()
    TextFieldGotFocus txtHomePhone
End Sub

Private Sub txtLastName_Change()

On Error GoTo ErrorTrap

    
    Call ChangedFieldsPersonal

    Exit Sub
ErrorTrap:
    EndProgram
End Sub

Private Sub txtLastName_GotFocus()
    TextFieldGotFocus txtLastName
End Sub

Private Sub txtLastName_LostFocus()
 
On Error GoTo ErrorTrap

       
    txtLastName = Trim(txtLastName)
'    txtLastName = DelSubstr(txtLastName, "'", False)
'    If Not (InStr(1, Trim(txtLastName), "McC") = 1 Or InStr(1, Trim(txtLastName), "McK") = 1) Then
'        txtLastName = StrConv(Trim(txtLastName), vbProperCase)
'    End If

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub txtMiddleName_Change()
    Call ChangedFieldsPersonal
End Sub


Private Sub txtFirstName_Change()
    Call ChangedFieldsPersonal
End Sub


Private Sub txtGender_Change()
    Call ChangedFieldsPersonal
End Sub


Private Sub txtDOB_Change()
    Call ChangedFieldsPersonal
End Sub


'Private Sub txtInfirmityLevel_Change()
'    Call ChangedFieldsPersonal
'End Sub


Private Sub chkActive_Click()

On Error GoTo ErrorTrap

    If DoNotTrigger Then Exit Sub
    
    Call ChangedFieldsPersonal
    
    If chkActive.value = vbChecked Then
        txtDOB.Enabled = True
        chkAnointed.Enabled = True
        cmbInfirmity.Enabled = True
        chkVisitingSpeaker.value = vbUnchecked
        chkVisitingSpeaker.Visible = False
        lblVisiting.Visible = False
        chkVisitingSpeaker.value = vbUnchecked
    Else
        txtDOB.text = ""
        cmbInfirmity.ListIndex = 0
        chkAnointed.value = vbUnchecked
        txtDOB.Enabled = False
        chkAnointed.Enabled = False
        cmbInfirmity.Enabled = False
        chkVisitingSpeaker.value = vbUnchecked
        chkVisitingSpeaker.Visible = True
        lblVisiting.Visible = True
    End If

    Exit Sub
ErrorTrap:
    EndProgram

    
End Sub
Private Sub chkAnointed_Click()
    Call ChangedFieldsPersonal
End Sub

Private Sub ChangedFieldsPersonal()

On Error GoTo ErrorTrap

    
    
    If UpdateRec Then
        cmdRefresh.Enabled = True
    End If
    
    cmdApply.Enabled = True
    
    cmdApply.Default = True
    

    Exit Sub
ErrorTrap:
    EndProgram
        
End Sub

Private Sub LockTextFieldsPersonal()

On Error GoTo ErrorTrap

    txtDOB.Locked = True
    txtFirstName.Locked = True
    cmbGender.Locked = True
    cmbInfirmity.Locked = True
    txtLastName.Locked = True
    txtMiddleName.Locked = True
    chkActive.Enabled = False
    chkAnointed.Enabled = False
    chkVisitingSpeaker.Enabled = False
    
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub LockTextFieldsFamily()

On Error GoTo ErrorTrap

    txtSpouse.Locked = True
    lstChildren.Enabled = True
    cmbAddFromList.Visible = False
    optgrpEditWhat.Visible = False
    lstParents.Enabled = False

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub LockTextFieldsAddress()

On Error GoTo ErrorTrap

    txtAddress1.Locked = True
    txtAddress2.Locked = True
    txtAddress3.Locked = True
    txtAddress4.Locked = True
    txtEmail.Locked = True
    txtMobile.Locked = True
    txtHomePhone.Locked = True
    txtPostcode.Locked = True
    cmbNames.Locked = True

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub UnlockTextFieldsAddress()

On Error GoTo ErrorTrap

    txtAddress1.Locked = False
    txtAddress2.Locked = False
    txtAddress3.Locked = False
    txtAddress4.Locked = False
    txtEmail.Locked = False
    txtMobile.Locked = False
    txtHomePhone.Locked = False
    txtPostcode.Locked = False
    cmbNames.Locked = False

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub UnlockTextFieldsPersonal()

On Error GoTo ErrorTrap

    txtDOB.Locked = False
    txtFirstName.Locked = False
    cmbGender.Locked = False
    cmbInfirmity.Locked = False
    txtLastName.Locked = False
    txtMiddleName.Locked = False
    chkActive.Enabled = True
    chkAnointed.Enabled = True
    chkVisitingSpeaker.Enabled = True

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub UnlockTextFieldsFamily()

On Error GoTo ErrorTrap


    
    lstChildren.Enabled = True
    cmbAddFromList.Visible = True
    optgrpEditWhat.Visible = True
    lblAddFamily.Visible = True
    
    Exit Sub
ErrorTrap:
    EndProgram
End Sub
Private Sub PopulateLinkedAddressNamesCombo()

On Error GoTo ErrorTrap

    cmbNames.Clear

    HandleListBox.PopulateListBox Me!cmbNames, _
        "SELECT DISTINCTROW tblNameAddress.ID, " & _
        "tblNameAddress.FirstName & ' ' & tblNameAddress.MiddleName, " & _
        "tblNameAddress.LastName " & _
        "FROM tblNameAddress " & _
        "WHERE Active = TRUE " & _
        " ORDER BY tblNameAddress.LastName, tblNameAddress.FirstName" _
        , CMSDB, 0, ", ", True, 2, 1
        
    cmbNames.AddItem "[NONE]", 0
    
    Exit Sub
ErrorTrap:
    EndProgram
End Sub


Private Sub GoToUpdateModePersonal()

On Error GoTo ErrorTrap

    'unlock text fields
    Call UnlockTextFieldsPersonal
    
    txtFirstName.SetFocus
    
    'enable/disable buttons and fields as appropriate
    lstNames.Enabled = False
    chkActiveOnly.Enabled = False
    cmdApply.Enabled = False
    cmdRefresh.Enabled = False
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = False
    cmdNew.Enabled = False
    cmdBrowse.Enabled = True
    cmdUpdate.Enabled = False
    txtSearch.Enabled = False
    cmdGo.Enabled = False
    
    If chkActiveOnly.value = vbUnchecked Then
        If cmbGender.ListIndex = 1 Then
            chkVisitingSpeaker.Visible = True
            lblVisiting.Visible = True
        Else
            chkVisitingSpeaker.Visible = False
            lblVisiting.Visible = False
        End If
    Else
        chkVisitingSpeaker.Visible = False
        lblVisiting.Visible = False
    End If
            
    
    UpdateRec = True
    NewRec = False


    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub GoToUpdateModeFamily()


On Error GoTo ErrorTrap


    'unlock text fields
    Call UnlockTextFieldsFamily
    txtSpouse.SetFocus
    
    'enable/disable buttons and fields as appropriate
    lstNames.Enabled = False
    chkActiveOnly.Enabled = False
    cmdApply.Enabled = False
    cmdRefresh.Enabled = False
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = False
    cmdNew.Enabled = False
    cmdBrowse.Enabled = True
    cmdUpdate.Enabled = False
    txtSearch.Enabled = False
    cmdGo.Enabled = False
    
    'Set up cmbAddFromList to show all possible spouses for person in lstNames
    optEditSpouse = True
    ListSpouses

    UpdateRec = True
    NewRec = False

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub GoToUpdateModeAddress()


On Error GoTo ErrorTrap


    'unlock text fields
    Call UnlockTextFieldsAddress
    txtAddress1.SetFocus
    
    'enable/disable buttons and fields as appropriate
    lstNames.Enabled = False
    chkActiveOnly.Enabled = False
    cmdApply.Enabled = False
    cmdRefresh.Enabled = False
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = False
    cmdNew.Enabled = False
    cmdBrowse.Enabled = True
    cmdUpdate.Enabled = False
    txtSearch.Enabled = False
    cmdGo.Enabled = False
    
    UpdateRec = True
    NewRec = False

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub ApplyChangesPersonal(SaveAndExit As Boolean)
Dim TheFirstName As String, TheLastName As String, TheMiddleName As String, CurrentID As Integer
Dim TempID As Long
On Error GoTo ErrorTrap

    If PersonalFieldsValidatedOK Then
     
        If MsgBox("Are you sure you want to save these changes?", vbYesNo + vbQuestion, _
                    AppName) = vbYes Then
                                                    
            If NewRec Then
                rstNameAddress.AddNew
            ElseIf UpdateRec Then
                rstNameAddress.Edit
            End If
            
            '
            'Updating tables this way (as opposed to using DB.EXECUTE) doesn't mind using
            ' a single apostrophe in the string. Hence DoubleUpSingleQuotes not required.
            '
            With rstNameAddress
                !FirstName = txtFirstName
                !LastName = txtLastName
                !MiddleName = txtMiddleName
                !GenderMF = IIf(cmbGender.ListIndex = 1, "M", "F")
                !DOB = IIf(txtDOB <> "", txtDOB, 0)
                ![InfirmityLevel0-6] = cmbInfirmity.ListIndex
                !Active = chkActive
                !Anointed = chkAnointed
                !LinkedAddressPerson = 0
                .Update
                
                TempID = !ID
                                
                rstNameAddress.Requery
                rstNameAddress2.Requery
                
                rstNameAddress.FindFirst "ID = " & TempID
                
                CurrentID = IIf(UpdateRec, !ID, rstNameAddress2!MaxID)
                
                
                rstNameAddress.FindFirst "ID = " & CurrentID
                                                
                If UpdateRec And chkActiveOnly.value = True And chkActive.value = False Then
             'find ID of next record in rstNameAddress so that this can be selected after update, since
             ' chkActiveOnly is True, and record now has chkActive = False
                    .MoveNext
                    If Not .EOF Then
                        HoldID = !ID
                        .MovePrevious
                    Else
                        .MoveFirst
                        HoldID = !ID
                        .MoveLast
                    End If
                    
                End If
                
                If chkActive.value = False Then
                    If chkVisitingSpeaker.value = vbUnchecked Then
                    ' person not in local cong. Delete from  visiting speaker
                        CongregationMember.DeleteVisitingSpeaker CLng(CurrentID)
                    Else
                        CongregationMember.MakePersonVisitingSpeaker 0, CLng(CurrentID)
                    End If
                Else
                    CongregationMember.DeleteVisitingSpeaker CLng(CurrentID)
                End If
                
                If UpdateRec Then
                    If chkActive.value = vbChecked Then
                        If Not PersonIsActive Then
                        'person been set to active
                            If Not CongregationMember.IsBaptised(!ID) And _
                                Not CongregationMember.IsAssociated(!ID, CLng(GlobalDefaultCong)) Then
                                'not bapt and not assoc - assoc them
                                    CongregationMember.AddPersonToRole GlobalDefaultCong, 5, 9, 55, CLng(HoldID)
                            End If
                        End If
                    End If
                End If
                
                If UpdateRec And _
                    PersonIsActive And _
                    chkActive.value = vbUnchecked Then
                    '
                    'Person has been set inactive, so remove them from school
                    ' and cut off their publisher dates.
                    '
                    SetPersonInactive CLng(CurrentID)
                End If
                
                PersonIsActive = chkActive.value
                
            End With
                
            rstNameAddress.FindFirst "ID = " & CLng(HoldID)
            
            If NewRec Then
                PopulateLinkedAddressNamesCombo 'refresh
            End If
            
            'now ensure spouse field is clear on Family Tab if this is a newly added individual
            If NewRec Then
                txtSpouse = ""
            End If
            
            'Associate person with congregation if new rec and active...
            If NewRec And chkActive.value = vbChecked Then
                CongregationMember.AddPersonToRole GlobalDefaultCong, 5, 9, 55, CLng(CurrentID)
            End If
                        
            If chkActiveOnly.value = vbChecked Then
                PopulateNamesList ShowActiveOnly:=True
            Else
                PopulateNamesList ShowActiveOnly:=False
            End If
            
            'if a new record has just been added, set lstNames to this new record. This is derived from rstNameAddress2,
            ' which gives MaxID - the maximum ID in rstNameAddress. This is the last one added.
            If NewRec Then
                If chkActiveOnly.value = vbChecked And chkActive.value = vbUnchecked Then
                    lstNames.ListIndex = -1
                Else
                    HandleListBox.SelectItem frmPersonalDetails!lstNames, CLng(CurrentID)
                    HoldID = rstNameAddress2!MaxID
                End If
            Else
                HandleListBox.SelectItem frmPersonalDetails!lstNames, CLng(HoldID)
            End If
            
            Select Case tabPeopleTabs.Tab
                Case 0: GoToBrowseModePersonal
                Case 1: GoToBrowseModeFamily
                Case 2: GoToBrowseModeRoles
            End Select
                       
            tabPeopleTabs.TabEnabled(1) = True
            tabPeopleTabs.TabEnabled(2) = True
            
            '
            'Person has been added to the system, so there's at least one person in it.
            ' Enable various controls if system previously had no people in it....
            '
            If Not SystemContainsPeople Then
                GlobalParms.Save "PeopleAreInTheSystem", "TrueFalse", True
                SystemContainsPeople = True
                CheckIfBrothersInDB
                frmMainMenu.EnforceSecurity
            End If
             
            
            '
            'Update the cache of names
            '
            CongregationMember.LoadNamesToCache
                       
            SaveAndExit = True
            
            NewRec = False
            UpdateRec = False
                        
            If FormIsOpen("frmCongStats") Then
                frmCongStats.UpdateCongStatsForm
                HandleListBox.Requery frmCongStats.cmbNames, True
            End If
      
        Else
            SaveAndExit = False
        End If
    Else
        SaveAndExit = False
    End If

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub MoveToNextRecPersonal()

On Error GoTo ErrorTrap

   
    HandleListBox.SelectItem frmPersonalDetails!lstNames, CLng(HoldID)
    rstNameAddress.FindFirst "ID = " & HoldID
    Call PopulateFieldsPersonal
    

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub GoToBrowseModePersonal()
'Repopulate text fields based on selection in lstNames

On Error GoTo ErrorTrap


    Call PopulateFieldsPersonal
    
    txtSearch.Enabled = True
    lstNames.Enabled = True
    'Me.Show
    'lstNames.SetFocus
    chkActiveOnly.Enabled = True
    
    If Not rstNameAddress.BOF Then
        HandleListBox.SelectItem frmPersonalDetails!lstNames, rstNameAddress!ID
    End If
    
' enable/disable buttons and fields as appropriate
    cmdApply.Enabled = False
    cmdUpdate.Enabled = True
    cmdRefresh.Enabled = False
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = True
    cmdNew.Enabled = True
    cmdBrowse.Enabled = False
    cmdGo.Enabled = True
    
    If chkActiveOnly.value = vbUnchecked Then
        If cmbGender.ListIndex = 1 Then
            chkVisitingSpeaker.Visible = True
            lblVisiting.Visible = True
        Else
            chkVisitingSpeaker.Visible = False
            lblVisiting.Visible = False
        End If
    Else
        chkVisitingSpeaker.Visible = False
        lblVisiting.Visible = False
    End If
        

'lock text fields
    Call LockTextFieldsPersonal

    
    NewRec = False
    UpdateRec = False


    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub GoToBrowseModeFamily()

On Error GoTo ErrorTrap

'Repopulate text fields based on selection in lstNames

    Call PopulateFieldsFamily
    
    txtSearch.Enabled = True
    lstNames.Enabled = True
    'lstNames.SetFocus
    chkActiveOnly.Enabled = True
    
    HandleListBox.SelectItem frmPersonalDetails!lstNames, rstNameAddress!ID
    
' enable/disable buttons and fields as appropriate
    cmdApply.Enabled = False
    cmdUpdate.Enabled = True
    cmdRefresh.Enabled = False
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = True
    cmdNew.Enabled = False
    cmdBrowse.Enabled = False
    cmdGo.Enabled = True

'lock text fields
    Call LockTextFieldsFamily

    
    NewRec = False
    UpdateRec = False


    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub GoToBrowseModeAddress()

On Error GoTo ErrorTrap

'Repopulate text fields based on selection in lstNames

    NewRec = False
    UpdateRec = False
    
    Call PopulateFieldsAddress
    
    txtSearch.Enabled = True
    lstNames.Enabled = True
    'lstNames.SetFocus
    chkActiveOnly.Enabled = True
    
    HandleListBox.SelectItem frmPersonalDetails!lstNames, rstNameAddress!ID
    
' enable/disable buttons and fields as appropriate
    cmdApply.Enabled = False
    cmdUpdate.Enabled = True
    cmdRefresh.Enabled = False
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = False
    cmdNew.Enabled = False
    cmdBrowse.Enabled = False
    cmdGo.Enabled = True

'lock text fields
    Call LockTextFieldsAddress

    


    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub GoToBrowseModeRoles()
'Repopulate text fields based on selection in lstNames

On Error GoTo ErrorTrap
 
    txtSearch.Enabled = True
    lstNames.Enabled = True
    chkActiveOnly.Enabled = True
    
'    HandleListBox.SelectItem frmPersonalDetails!lstNames, rstNameAddress!ID
    
' enable/disable buttons and fields as appropriate
    cmdApply.Enabled = False
    cmdUpdate.Enabled = False
    cmdRefresh.Enabled = False
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = False
    cmdNew.Enabled = False
    cmdBrowse.Enabled = False
    cmdGo.Enabled = True
    cmdSuspendDates.Enabled = True
    
    NewRec = False
    UpdateRec = False


    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub



Private Sub txtMiddleName_GotFocus()
    TextFieldGotFocus txtMiddleName
End Sub

Private Sub txtMiddleName_LostFocus()

On Error GoTo ErrorTrap

    ' if in browse or new mode, check field on exit from it
    
    If NewRec Or UpdateRec Then
        
        txtMiddleName = Trim(txtMiddleName)
'        If Not (InStr(1, Trim(txtMiddleName), "McC") = 1 Or InStr(1, Trim(txtMiddleName), "McK") = 1) Then
'            txtMiddleName = StrConv(Trim(txtMiddleName), vbProperCase)
'        End If
        
       ' txtMiddleName = DelSubstr(txtMiddleName, "'", False)
        
        If Len(txtMiddleName) > 100 Then
            If MsgBox("Middle name should be no greater than 100 characters long. " & _
                      "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
                txtMiddleName.SetFocus
            Else
                'field is invalid, so forget it and go to browse-mode since user clicked Cancel.
                GoToBrowseModePersonal
            End If
        End If
    End If

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Function PersonalFieldsValidatedOK() As Boolean
        
Dim rstNameCheck As Recordset, Query As String
Dim TheFirstName As String, TheLastName As String, TheMiddleName As String


On Error GoTo ErrorTrap

    PersonalFieldsValidatedOK = True
        
    'Validate txtFirstName
        
    If Len(txtFirstName) > 100 Or Len(txtFirstName) = 0 Then
        PersonalFieldsValidatedOK = False
        If MsgBox("First name should be between 1 and 100 characters long. " & _
                  "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
            
        Else
            'field is invalid, so forget it and go to browse-mode since user clicked Cancel.
            GoToBrowseModePersonal
        End If
        txtFirstName.SetFocus
        Exit Function
    Else
        PersonalFieldsValidatedOK = True
    End If
    
    'Validate txtLastName

    If Len(txtLastName) > 100 Or Len(txtLastName) = 0 Then
        PersonalFieldsValidatedOK = False
        If MsgBox("Last name should be between 1 and 100 characters long. " & _
                  "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
            
        Else
            'field is invalid, so forget it and go to browse-mode since user clicked Cancel.
            GoToBrowseModePersonal
        End If
        txtLastName.SetFocus
        Exit Function
    Else
        PersonalFieldsValidatedOK = True
    End If
    
    
    'Validate txtMiddleName

    If Len(txtMiddleName) > 100 Then
        PersonalFieldsValidatedOK = False
        If MsgBox("Middle name should be no greater than 100 characters long. " & _
                  "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
            
        Else
            'field is invalid, so forget it and go to browse-mode since user clicked Cancel.
            GoToBrowseModePersonal
        End If
        txtMiddleName.SetFocus
        Exit Function
    Else
        PersonalFieldsValidatedOK = True
    End If

       
    'Check if person's name already appears in tblNameAddress - only if it's changed
    
    If StoredFirstName <> txtFirstName.text Or _
       StoredMiddleName <> txtMiddleName.text Or _
       StoredLastName <> txtLastName.text Then
       
        TheFirstName = DoubleUpSingleQuotes(txtFirstName)
        TheMiddleName = DoubleUpSingleQuotes(txtMiddleName)
        TheLastName = DoubleUpSingleQuotes(txtLastName)
        
        'this IF is done simply because, if txtMiddleName is blank the SQL wouldn't work
        If txtMiddleName = "" Or IsNull(txtMiddleName) Then
            Query = "SELECT * FROM tblNameAddress WHERE FirstName = '" & TheFirstName & "'" & " And LastName = '" & TheLastName & "'" & " AND (IsNull(MiddleName) OR Len(MiddleName)) = 0"
        Else
            Query = "SELECT * FROM tblNameAddress WHERE FirstName = '" & TheFirstName & "'" & " And LastName = '" & TheLastName & "'" & " AND MiddleName = '" & TheMiddleName & "'"
        End If
    
        Set rstNameCheck = CMSDB.OpenRecordset(Query, dbOpenDynaset)
    
        
        If Not rstNameCheck.BOF Then
            PersonalFieldsValidatedOK = False
            If MsgBox(txtFirstName & " " & txtMiddleName & " " & txtLastName & " already exists." & _
                      " Try to make it unique by altering the Middle Name, for example." & _
                      " Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
            Else
                'field is invalid, so forget it and go to browse-mode since user clicked Cancel.
                GoToBrowseModePersonal
            End If
            txtMiddleName.SetFocus
            rstNameCheck.Close
            Exit Function
        Else
            rstNameCheck.Close
            PersonalFieldsValidatedOK = True
        End If
    End If
    
    If cmbGender.ListIndex = -1 Then
        PersonalFieldsValidatedOK = False
        If MsgBox("Select gender. " & _
                  "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
        Else
            'field is invalid, so forget it and go to browse-mode since user clicked Cancel.
            GoToBrowseModePersonal
        End If
        cmbGender.SetFocus
        Exit Function
    Else
        If chkVisitingSpeaker.value = vbChecked Then
            If cmbGender.ListIndex = 0 Then
                PersonalFieldsValidatedOK = False
                If MsgBox("Only males can be visiting speakers. " & _
                          "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
                Else
                    'field is invalid, so forget it and go to browse-mode since user clicked Cancel.
                    GoToBrowseModePersonal
                End If
                cmbGender.SetFocus
                Exit Function
            Else
                PersonalFieldsValidatedOK = True
            End If
        End If
    End If
    
    'Validate txtInfirmityLevel
    
'    txtInfirmityLevel = Trim(txtInfirmityLevel)
    If cmbInfirmity.ListIndex = -1 Then
        cmbInfirmity.ListIndex = 0
    End If
    
     
     'Validate txtDOB

    If chkActive.value = vbChecked Then
        If Not IsDate(txtDOB) Then
            PersonalFieldsValidatedOK = False
            If MsgBox("Date is not valid. " & _
                      "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
            Else
                'field is invalid, so forget entire changes and go to browse-mode since user clicked Cancel.
                GoToBrowseModePersonal
            End If
            txtDOB.SetFocus
            Exit Function
        Else
            PersonalFieldsValidatedOK = True
        End If
        
        'Validate txtDOB - has the person been born in the future???
        If CDate(txtDOB) > date Then
            PersonalFieldsValidatedOK = False
            If MsgBox("Date of birth is in the future! " & _
                      "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
            Else
                'field is invalid, so forget entire changes and go to browse-mode since user clicked Cancel.
                GoToBrowseModePersonal
            End If
            txtDOB.SetFocus
            Exit Function
        Else
            PersonalFieldsValidatedOK = True
        End If
    End If
        
    Exit Function
ErrorTrap:
    EndProgram
    
End Function

Private Function AddressFieldsValidatedOK() As Boolean
        
Dim rstRecSet As Recordset, SQLStr As String
        
On Error GoTo ErrorTrap

    AddressFieldsValidatedOK = True
    
    txtPostcode.text = StrConv(txtPostcode.text, vbUpperCase)
    txtAddress1.text = StrConv(txtAddress1.text, vbProperCase)
    txtAddress2.text = StrConv(txtAddress2.text, vbProperCase)
    txtAddress3.text = StrConv(txtAddress3.text, vbProperCase)
    txtAddress4.text = StrConv(txtAddress4.text, vbProperCase)
        
    'Validate email address
    If Not ValidateEmailAddress(txtEmail) Then
        AddressFieldsValidatedOK = False
        If MsgBox("Invalid email address. " & _
                  "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
        Else
            'field is invalid, so forget it and go to browse-mode since user clicked Cancel.
            GoToBrowseModeAddress
        End If
        txtEmail.SetFocus
        SendKeys "{Home}+{End}"
        Exit Function
    Else
        AddressFieldsValidatedOK = True
    End If
    
    'check linked combo set
    If cmbNames.ListIndex = -1 Then
        AddressFieldsValidatedOK = False
        If MsgBox("Invalid name selected. " & _
                  "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
        Else
            'field is invalid, so forget it and go to browse-mode since user clicked Cancel.
            GoToBrowseModeAddress
        End If
        cmbNames.SetFocus
        Exit Function
    Else
        AddressFieldsValidatedOK = True
    End If
    
    'check not linking address to same person
    If cmbNames.ItemData(cmbNames.ListIndex) = lstNames.ItemData(lstNames.ListIndex) Then
        AddressFieldsValidatedOK = False
        If MsgBox("Cannot link person to their own address. " & _
                  "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
        Else
            'field is invalid, so forget it and go to browse-mode since user clicked Cancel.
            GoToBrowseModeAddress
        End If
        cmbNames.SetFocus
        Exit Function
    Else
        AddressFieldsValidatedOK = True
    End If
    
    'check person to whom linking is not already linked to someone else
    If cmbNames.ListIndex > 0 Then
        SQLStr = "SELECT ID FROM tblNameAddress " & _
                " WHERE ID = " & cmbNames.ItemData(cmbNames.ListIndex) & _
                " AND LinkedAddressPerson <> 0 " & _
                " AND LinkedAddressPerson IS NOT NULL "

        Set rstRecSet = CMSDB.OpenRecordset(SQLStr, dbOpenDynaset)
    
        
        If Not rstRecSet.BOF Then
            AddressFieldsValidatedOK = False
            If MsgBox("The person to whom the address is being linked is already linked to another address." & _
                      " Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
            Else
                'field is invalid, so forget it and go to browse-mode since user clicked Cancel.
                GoToBrowseModeAddress
            End If
            cmbNames.SetFocus
            rstRecSet.Close
            Exit Function
        Else
            rstRecSet.Close
            AddressFieldsValidatedOK = True
        End If
    End If

    Exit Function
ErrorTrap:
    EndProgram
    
End Function


Private Sub DeleteRecPersonal()

On Error GoTo ErrorTrap

    If rstNameAddress.BOF Then Exit Sub
    
    If MsgBox("Are you sure you want to delete " & _
              CongregationMember.FirstAndLastName(rstNameAddress!ID) & _
              " from CMS? ALL corresponding data will also be deleted!", vbYesNo + vbQuestion, _
                AppName) = vbYes Then
        If MsgBox("ARE YOU ABSOLUTELY SURE? ALL corresponding data will also be deleted!", vbYesNo + vbQuestion, _
                    AppName) = vbYes Then
        
        'Delete corresponding data on related tables
            If DeleteSomeRows("tblIDWeightings", "ID = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblMarriage", "ID = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblMarriage", "Spouse = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblChildren", "Parent = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblChildren", "Child = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblTaskAndPerson", "Person = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblTaskPersonSuspendDates", "Person = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblTMSCounselPoints", "StudentID = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblBaptismDates", "PersonID = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblBookGroupMembers", "PersonID = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblEldersAndServants", "PersonID = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblIrregularPubs", "PersonID = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblMinReports", "PersonID = ", rstNameAddress!ID) And _
               DeleteMissingReportsForPerson(rstNameAddress!ID) And _
               DeleteSomeRows("tblPublisherDates", "PersonID = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblRegPioDates", "PersonID = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblAuxPioDates", "PersonID = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblSpecPioDates", "PersonID = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblPubRecCardRowPrinted", "PersonID = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblVisitingSpeakers", "PersonID = ", rstNameAddress!ID) And _
               DeleteSomeRows("tblIndividualSPAMWeightings", "PersonID = ", rstNameAddress!ID) Then
                With rstNameAddress
                'find ID of next record in rstNameAddress so that this can be selected after delete
                    .MoveNext
                    If Not .EOF Then
                        HoldID = !ID
                        .MovePrevious
                    Else
                        .MoveFirst
                        HoldID = !ID
                        .MoveLast
                    End If
                    
                    .Delete
                    .Requery
                        
                    If .BOF Then 'no people in the system.
                        tabPeopleTabs.TabEnabled(1) = False
                        tabPeopleTabs.TabEnabled(2) = False
                        GlobalParms.Save "PeopleAreInTheSystem", "TrueFalse", False
                        CheckIfBrothersInDB
                    End If
                End With
                
                '
                'Update the cache of names
                '
                CongregationMember.LoadNamesToCache
                
                PopulateLinkedAddressNamesCombo
                
                rstNameAddress2.Requery
                HandleListBox.Requery lstNames, True
                lstNames.Enabled = True
                lstNames.SetFocus
                cmdApply.Enabled = False
                cmdNew.Enabled = True
                cmdDelete.Enabled = True
                cmdUpdate.Enabled = True
                cmdRefresh.Enabled = False
                cmdOK.Enabled = True
                cmdCancel.Enabled = True
                cmdBrowse.Enabled = False
                
        'lock text fields
                Call LockTextFieldsPersonal
                
                MoveToNextRecPersonal
                
                NewRec = False
                UpdateRec = False
              
            End If
    
        End If
    End If

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub GoToNewModePersonal()

On Error GoTo ErrorTrap

    lstNames.Enabled = False
    chkActiveOnly.Enabled = False
    
    txtFirstName = ""
    txtLastName = ""
    txtMiddleName = ""
    
    '
    'Store the name fields for use in checking whether user has changed them
    '
    StoredFirstName = ""
    StoredMiddleName = ""
    StoredLastName = ""

    cmbGender.ListIndex = -1
    cmbGender.text = ""
    txtFirstName = ""
    txtDOB = ""
    cmbInfirmity.ListIndex = 0
    chkActive = vbChecked
    cmdBrowse.Enabled = True
    cmdApply.Enabled = False
    cmdUpdate.Enabled = False
    cmdRefresh.Enabled = False
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = False
    txtSearch.Enabled = False
    cmdGo.Enabled = False
    
    NewRec = True
    UpdateRec = False
    
'Unlock text fields
    Call UnlockTextFieldsPersonal
    
    txtFirstName.SetFocus
    cmdNew.Enabled = False


    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub RefreshPersonal()

On Error GoTo ErrorTrap

    
    Call PopulateFieldsPersonal
    txtFirstName.SetFocus
    cmdRefresh.Enabled = False
    cmdApply.Enabled = False
    
    HandleListBox.SelectItem lstNames, rstNameAddress!ID
    
    NewRec = False

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

'Private Sub RefreshResp()
'
'On Error GoTo ErrorTrap
'
'
'    cmdRefresh.Enabled = False
'    cmdApply.Enabled = False
'
'    NewRec = False
'
'    Exit Sub
'ErrorTrap:
'    EndProgram
'
'
'End Sub



Private Sub SetUpPersonalTab()

On Error GoTo ErrorTrap

    If rstNameAddress.BOF Then Exit Sub
    
    Call PopulateFieldsPersonal
    
    PopulateLinkedAddressNamesCombo
    
       
    ' Find the record that matches the list box.
    rstNameAddress.MoveFirst
    HandleListBox.SelectItem frmPersonalDetails!lstNames, rstNameAddress!ID
       
    'frmPersonalDetails!lstNames.SetFocus
    
    HoldID = rstNameAddress!ID

    ' enable/disable buttons and fields as appropriate
    cmdApply.Enabled = False
    cmdUpdate.Enabled = True
    cmdRefresh.Enabled = False
    cmdOK.Enabled = True
    cmdCancel.Enabled = True
    cmdDelete.Enabled = True
    cmdNew.Enabled = True
    cmdBrowse.Enabled = False
    txtSearch.Enabled = True
    
    If Len(txtSearch) = 0 Or IsNull(txtSearch) Then
        cmdGo.Enabled = False
    Else
        cmdGo.Enabled = True
    End If
    
    If chkActiveOnly.value = vbUnchecked Then
        If cmbGender.ListIndex = 1 Then
            chkVisitingSpeaker.Visible = True
            lblVisiting.Visible = True
        Else
            chkVisitingSpeaker.Visible = False
            lblVisiting.Visible = False
        End If
    Else
        chkVisitingSpeaker.Visible = False
        lblVisiting.Visible = False
    End If
            

'lock text fields
    Call LockTextFieldsPersonal
    
    UpdateRec = False
    NewRec = False

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub SetUpFamilyTab()

On Error GoTo ErrorTrap

    If rstNameAddress.BOF Then Exit Sub
    
       
' Find the record that matches the list box.
    rstNameAddress.MoveFirst
    HandleListBox.SelectItem lstNames, rstNameAddress!ID
    
    Call PopulateFieldsFamily
    'lstNames.SetFocus
    
    HoldID = rstNameAddress!ID

'lock text fields
    Call LockTextFieldsFamily
    
    optEditSpouse = True
    optgrpEditWhat.Visible = False
    cmbAddFromList.Visible = False
    lblAddFamily.Visible = False

    Exit Sub
ErrorTrap:
    EndProgram
    
    
End Sub
Private Sub SetUpAddressTab()

On Error GoTo ErrorTrap

    If rstNameAddress.BOF Then Exit Sub
    
       
' Find the record that matches the list box.
    rstNameAddress.MoveFirst
    HandleListBox.SelectItem lstNames, rstNameAddress!ID
        
    Call PopulateFieldsAddress
    
    HoldID = rstNameAddress!ID

'lock text fields
    Call LockTextFieldsAddress
    
    
    Exit Sub
ErrorTrap:
    EndProgram
    
    
End Sub



Private Sub SelectFromLstNamesPersonal()

On Error GoTo ErrorTrap

    ' Find the record that matches the control.
    rstNameAddress.FindFirst "[ID] = " & lstNames.ItemData(lstNames.ListIndex)
    Call PopulateFieldsPersonal
    cmdApply.Enabled = False
    HoldID = rstNameAddress!ID
    TheSelectedPerson = CLng(HoldID)

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub

Private Sub SelectFromLstNamesFamily()

On Error GoTo ErrorTrap


    Call PopulateFieldsFamily

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub SelectFromLstNamesAddress()

On Error GoTo ErrorTrap

    If rstNameAddress!LinkedAddressPerson > 0 Then
        HandleListBox.SelectItem cmbNames, rstNameAddress!LinkedAddressPerson
    Else
        If cmbNames.ListCount > 0 Then
            cmbNames.ListIndex = 0
        End If
    End If
    
    Call PopulateFieldsAddress

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub txtMobile_Change()
    If UpdateRec Then
        cmdApply.Enabled = True
    End If

End Sub

Private Sub txtMobile_GotFocus()
    TextFieldGotFocus txtMobile
End Sub

Private Sub txtPostcode_Change()
    If UpdateRec Then
        cmdApply.Enabled = True
    End If

End Sub

Private Sub txtPostcode_GotFocus()
    TextFieldGotFocus txtPostcode
End Sub

Private Sub txtSearch_Change()

On Error GoTo ErrorTrap
    
    cmdGo.Enabled = True
    cmdGo.Default = True
    cmdNext.Enabled = False
    cmdPrev.Enabled = False
    SearchActive = False
    
    Exit Sub
ErrorTrap:
    EndProgram
End Sub


Private Sub PopulateFieldsFamily()
'populate the fields on the Family tab...

On Error GoTo ErrorTrap


Dim MarriageSQL As String, ChildrenSQL As String, ParentSQL As String
    
'Find the spouse based on lstNames
    MarriageSQL = "SELECT tblNameAddress.ID, " & _
                         "tblNameAddress.FirstName, " & _
                         "tblNameAddress.MiddleName, " & _
                         "tblNameAddress.LastName " & _
                  "FROM tblNameAddress " & _
                  "WHERE tblNameAddress.ID = (SELECT tblMarriage.Spouse " & _
                                             "FROM tblMarriage " & _
                                             "WHERE tblMarriage.ID = " & lstNames.ItemData(lstNames.ListIndex) & ")"
                                            
    Set rstMarriage = CMSDB.OpenRecordset(MarriageSQL, dbOpenDynaset)
        
    With rstMarriage
        
        If Not .BOF Then
            If IsNull(!MiddleName) Or Len(!MiddleName) = 0 Then
                txtSpouse = !FirstName & " " & !LastName
            Else
                txtSpouse = !FirstName & " " & !MiddleName & " " & !LastName
            End If
            TheNewSpouse = !ID
        Else
            txtSpouse = ""
        End If
                
        .Close
    End With
    
'Find the children, based on lstNames
    
    ChildrenSQL = "SELECT tblNameAddress.ID, " & _
                         "tblNameAddress.FirstName," & _
                         "tblNameAddress.MiddleName," & _
                         "tblNameAddress.LastName " & _
                  "FROM tblNameAddress " & _
                  "WHERE tblNameAddress.ID IN (SELECT tblChildren.Child " & _
                                              "FROM tblChildren " & _
                                              "WHERE tblChildren.Parent = " & lstNames.ItemData(lstNames.ListIndex) & ")"
    
    HandleListBox.PopulateListBox frmPersonalDetails!lstChildren, ChildrenSQL, CMSDB, 0, " ", True, 1, 2, 3
    
'Find the parents, based on lstNames

    ParentSQL = "SELECT ID, tblNameAddress.FirstName," & _
                         "tblNameAddress.MiddleName," & _
                         "tblNameAddress.LastName " & _
                  "FROM tblNameAddress " & _
                  "WHERE tblNameAddress.ID IN (SELECT tblChildren.Parent " & _
                                              "FROM tblChildren " & _
                                              "WHERE tblChildren.Child = " & lstNames.ItemData(lstNames.ListIndex) & ")"
        
    HandleListBox.PopulateListBox frmPersonalDetails!lstParents, ParentSQL, CMSDB, 0, " ", True, 1, 2, 3

    

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub PopulateFieldsAddress()
'populate the fields on the address tab...

Dim ThePerson As Long

On Error GoTo ErrorTrap

    ThePerson = lstNames.ItemData(lstNames.ListIndex)

    DoNotTriggerAddressStuff = True
    
    If cmbNames.ListCount > 0 Then
        If Not IsNull(rstNameAddress!LinkedAddressPerson) Then
            If rstNameAddress!LinkedAddressPerson > 0 Then
                If Not UpdateRec Then
                    HandleListBox.SelectItem cmbNames, rstNameAddress!LinkedAddressPerson
                End If
            Else
                If Not UpdateRec Then
                    cmbNames.ListIndex = 0
                End If
            End If
        Else
            If Not UpdateRec Then
                cmbNames.ListIndex = 0
            End If
        End If
    End If

    DoNotTriggerAddressStuff = False
    
    If cmbNames.ListIndex > 0 Then
        ShadeOutAddressFields True
    Else
        ShadeOutAddressFields False
    End If

    txtAddress1 = CongregationMember.Address1(ThePerson)
    txtAddress2 = CongregationMember.Address2(ThePerson)
    txtAddress3 = CongregationMember.Address3(ThePerson)
    txtAddress4 = CongregationMember.Address4(ThePerson)
    txtPostcode = CongregationMember.PostCode(ThePerson)
    txtHomePhone = CongregationMember.HomePhone(ThePerson)
    
    txtMobile = CongregationMember.MobilePhone(ThePerson)
    txtEmail = CongregationMember.Email(ThePerson)

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub ShadeOutAddressFields(ShadeThem As Boolean)

    If ShadeThem Then
        txtAddress1.BackColor = RGB(225, 225, 225) 'grey
        txtAddress2.BackColor = RGB(225, 225, 225) 'grey
        txtAddress3.BackColor = RGB(225, 225, 225) 'grey
        txtAddress4.BackColor = RGB(225, 225, 225) 'grey
        txtPostcode.BackColor = RGB(225, 225, 225) 'grey
        txtHomePhone.BackColor = RGB(225, 225, 225) 'grey
    Else
        txtAddress1.BackColor = RGB(255, 255, 255) 'white
        txtAddress2.BackColor = RGB(255, 255, 255) 'white
        txtAddress3.BackColor = RGB(255, 255, 255) 'white
        txtAddress4.BackColor = RGB(255, 255, 255) 'white
        txtPostcode.BackColor = RGB(255, 255, 255) 'white
        txtHomePhone.BackColor = RGB(255, 255, 255) 'white
    End If



End Sub

Private Sub txtSearch_GotFocus()
    TextFieldGotFocus txtSearch
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    'MsgBox "KEY: " & KeyCode, vbOKOnly
'        MsgBox "KEY: " & KeyAscii, vbOKOnly
 '   If KeyCode = 13 Then
    'ENTER key pressed
 '       GoSearchForName
 '   End If

End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 13 Then
    'ENTER key pressed
'        GoSearchForName
'    End If
End Sub

Private Sub txtSpouse_GotFocus()

On Error GoTo ErrorTrap

    ListSpouses
    optEditSpouse = True

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub ApplyChangesFamily(SaveAndExit As Boolean)
Dim TheCurrentSpouse As Integer

On Error GoTo ErrorTrap


    If FamilyFieldsValidatedOK Then
     
        If MsgBox("Are you sure you want to save these changes?", vbYesNo + vbQuestion, _
                    AppName) = vbYes Then
            
            If ReplaceSpouse Or NewSpouse Then
                
                TheNewSpouse = cmbAddFromList.ItemData(cmbAddFromList.ListIndex)
                
                Set rstMarriage = CMSDB.OpenRecordset("SELECT * FROM tblMarriage", dbOpenDynaset)
                With rstMarriage
                
                If ReplaceSpouse Then
                'Delete the link between party A and B on tblMarriage
                    .FindFirst ("ID = " & lstNames.ItemData(lstNames.ListIndex))
                    TheCurrentSpouse = !Spouse
                    .Delete
                    .Requery
                'Now delete the link between party B and A on tblMarriage
                    .FindFirst ("ID = " & TheCurrentSpouse)
                    .Delete
                    .Requery
                End If
                
                'now insert link between party A and C (where C is the new spouse to add)
                .AddNew
                !ID = lstNames.ItemData(lstNames.ListIndex)
                !Spouse = TheNewSpouse
                .Update
                'now insert link between party C and A (where C is the new spouse to add)
                .AddNew
                !ID = TheNewSpouse
                !Spouse = lstNames.ItemData(lstNames.ListIndex)
                .Update
                .Requery
                .Close
                End With
            ElseIf NewChild Then
                Set rstChildren = CMSDB.OpenRecordset("SELECT * FROM tblChildren", dbOpenDynaset)
                With rstChildren
                
                'now insert link between party A and B (where B is the new child to add)
                .AddNew
                !Parent = lstNames.ItemData(lstNames.ListIndex)
                !Child = cmbAddFromList.ItemData(cmbAddFromList.ListIndex)
                .Update
                .Requery
                .Close
                End With
                
            End If
            
            Select Case tabPeopleTabs.Tab
                Case 0: GoToBrowseModePersonal
                Case 1: GoToBrowseModeFamily
            End Select
            
            cmbAddFromList.ListIndex = -1
            
            SaveAndExit = True
            
            NewRec = False
            UpdateRec = False
            NewSpouse = False
            ReplaceSpouse = False
            NewChild = False
        Else
            SaveAndExit = False
        End If
    Else
        SaveAndExit = False
    End If

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub
Private Sub ApplyChangesAddress(SaveAndExit As Boolean)

On Error GoTo ErrorTrap


    If MsgBox("Are you sure you want to save these changes?", vbYesNo + vbQuestion, _
                AppName) = vbYes Then
                
        If Not AddressFieldsValidatedOK Then
            SaveAndExit = False
            Exit Sub
        End If
        
        With rstNameAddress
        
        .Edit
        
        If cmbNames.ListIndex > 0 Then 'linked address
            !LinkedAddressPerson = cmbNames.ItemData(cmbNames.ListIndex)
            !Address1 = ""
            !Address2 = ""
            !Address3 = ""
            !Address4 = ""
            !PostCode = ""
            !HomePhone = ""
        Else
            !LinkedAddressPerson = 0
            !Address1 = Trim(txtAddress1.text)
            !Address2 = Trim(txtAddress2.text)
            !Address3 = Trim(txtAddress3.text)
            !Address4 = Trim(txtAddress4.text)
            !PostCode = Trim(txtPostcode.text)
            !HomePhone = Trim(txtHomePhone.text)
        End If
        
        !Email = Trim(txtEmail.text)
        !MobilePhone = Trim(txtMobile.text)
        
        .Update
        
        End With
        
        Select Case tabPeopleTabs.Tab
            Case 0: GoToBrowseModePersonal
            Case 1: GoToBrowseModeFamily
            Case 3: GoToBrowseModeAddress
        End Select
        
        
        SaveAndExit = True
        
        NewRec = False
        UpdateRec = False
        NewSpouse = False
        ReplaceSpouse = False
        NewChild = False
    Else
        SaveAndExit = False
    End If

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Function FamilyFieldsValidatedOK() As Boolean

'Check that a name has been selected on cmbAddFromList

On Error GoTo ErrorTrap

    If cmbAddFromList.ListIndex = -1 Then
        FamilyFieldsValidatedOK = False
        If MsgBox("Please select a name." & _
                  "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
                    cmbAddFromList.SetFocus
        Else
            'field is invalid, so forget it and go to browse-mode since user clicked Cancel.
            GoToBrowseModeFamily
        End If
        Exit Function
    Else
        FamilyFieldsValidatedOK = True
    End If


    Exit Function
ErrorTrap:
    EndProgram
    
    
End Function

Private Sub DeleteRecFamily()
Dim TheCurrentSpouse As Integer

On Error GoTo ErrorTrap


    
    If ReplaceSpouse And txtSpouse <> "" Then
        If MsgBox("Are you sure you want to break the link between " & _
                  txtFirstName & " " & txtMiddleName & " " & txtLastName & " and " & _
                  txtSpouse & "?", vbYesNo + vbQuestion, AppName) = vbYes Then
            Set rstMarriage = CMSDB.OpenRecordset("SELECT * FROM tblMarriage", dbOpenDynaset)
            With rstMarriage
            'Delete the link between party A and B on tblMarriage
            .FindFirst ("ID = " & lstNames.ItemData(lstNames.ListIndex))
            TheCurrentSpouse = !Spouse
            .Delete
            .Requery
            'Now delete the link between party B and A on tblMarriage
            .FindFirst ("ID = " & TheCurrentSpouse)
            .Delete
            .Requery
            .Close
            End With
        End If
    ElseIf NewChild And lstChildren.ListIndex > -1 Then
        If MsgBox("Are you sure you want to break the link between " & _
          txtFirstName & " " & txtMiddleName & " " & txtLastName & " and " & _
          lstChildren.text & "?", vbYesNo + vbQuestion, AppName) = vbYes Then

            Set rstChildren = CMSDB.OpenRecordset("SELECT * FROM tblChildren", dbOpenDynaset)
            With rstChildren
            'Delete the link between party A and B on tblChildren
            .FindFirst ("Child = " & lstChildren.ItemData(lstChildren.ListIndex))
            .Delete
            .Requery
            .Close
            End With
        End If
    Else
        MsgBox "Please select a Spouse or Child", vbOKOnly + vbInformation, AppName
    End If
    
    PopulateFieldsFamily


    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub SetUpResponsibilitiesTab()
       
' Find the record that matches the list box.

On Error GoTo ErrorTrap

    If rstNameAddress.BOF Then Exit Sub
    
    rstNameAddress.MoveFirst
    HandleListBox.SelectItem lstNames, rstNameAddress!ID
    
'    'lstNames.SetFocus
'
'    Call PopulateFieldsRoles(True)
'
    HoldID = rstNameAddress!ID
'
'    If cmbCongregation.ListIndex = -1 Then
'        cmdSuspendDates.Enabled = False
'    Else
'        cmdSuspendDates.Enabled = True
'    End If
    

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub



Private Sub EnforceSecurity()
    On Error GoTo ErrorTrap
    
    If Not AccessAllowed(frmPersonalDetails.Name, cmdSPAM.Name) Then
        cmdSPAM.Enabled = False
    End If
    If Not AccessAllowed(frmPersonalDetails.Name, cmdTMSStudents.Name) Then
        cmdTMSStudents.Enabled = False
    End If
    If Not AccessAllowed(frmPersonalDetails.Name, cmdPublicMeetingPersonnel.Name) Then
        cmdPublicMeetingPersonnel.Enabled = False
    End If
    If Not AccessAllowed(frmPersonalDetails.Name, cmdPublicMeetingPersonnel.Name) Then
        cmdServiceMtgs.Enabled = False
    End If

    Exit Sub
ErrorTrap:
    EndProgram

End Sub

Public Sub PersDtlsForm_EnforceSecurity()
    EnforceSecurity
End Sub

Public Sub SetUpNameAddressRecSets()

    On Error GoTo ErrorTrap
    
    Set rstNameAddress = CMSDB.OpenRecordset("SELECT * " & _
                                            "FROM tblNameAddress " & _
                                            "WHERE Active = TRUE " & _
                                            "ORDER BY LastName, FirstName" _
                                            , dbOpenDynaset)
                                            
    Set rstNameAddress2 = CMSDB.OpenRecordset("SELECT MAX(ID) as MaxID " & _
                                            "FROM tblNameAddress " _
                                            , dbOpenForwardOnly)
                                            
    Exit Sub
    
ErrorTrap:
    EndProgram

End Sub


Public Property Get FormPersonID() As Long
    FormPersonID = TheSelectedPerson
End Property

Public Property Let FormPersonID(ByVal vNewValue As Long)
    TheSelectedPerson = vNewValue
End Property

