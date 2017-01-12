VERSION 5.00
Begin VB.Form frmAddPerson 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " C.M.S. Add/Edit Person"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5265
   Icon            =   "frmAddPerson.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtEmail 
      Height          =   314
      Left            =   1365
      MaxLength       =   100
      TabIndex        =   6
      Top             =   2745
      Width           =   3729
   End
   Begin VB.TextBox txtMobile 
      Height          =   314
      Left            =   1365
      MaxLength       =   20
      TabIndex        =   4
      Top             =   2025
      Width           =   1920
   End
   Begin VB.TextBox txtHomePhone 
      Height          =   314
      Left            =   1365
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1650
      Width           =   1920
   End
   Begin VB.TextBox txtMobile2 
      Height          =   314
      Left            =   1365
      MaxLength       =   20
      TabIndex        =   5
      Top             =   2385
      Width           =   1920
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   330
      Left            =   4140
      TabIndex        =   8
      Top             =   3225
      Width           =   870
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   330
      Left            =   3210
      TabIndex        =   7
      Top             =   3225
      Width           =   870
   End
   Begin VB.TextBox txtLastName 
      Height          =   314
      Left            =   1365
      MaxLength       =   50
      TabIndex        =   2
      Top             =   1155
      Width           =   3729
   End
   Begin VB.TextBox txtMiddleName 
      Height          =   314
      Left            =   1365
      MaxLength       =   50
      TabIndex        =   1
      Top             =   795
      Width           =   3729
   End
   Begin VB.TextBox txtFirstName 
      Height          =   314
      Left            =   1365
      MaxLength       =   50
      TabIndex        =   0
      Top             =   435
      Width           =   3729
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Home Phone:"
      Height          =   240
      Left            =   135
      TabIndex        =   16
      Top             =   1680
      Width           =   1365
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Phone:"
      Height          =   240
      Left            =   135
      TabIndex        =   15
      Top             =   2070
      Width           =   1365
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Email Address:"
      Height          =   240
      Left            =   135
      TabIndex        =   14
      Top             =   2790
      Width           =   1365
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Phone 2:"
      Height          =   240
      Left            =   135
      TabIndex        =   13
      Top             =   2445
      Width           =   1365
   End
   Begin VB.Label lblCongName 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      Height          =   240
      Left            =   135
      TabIndex        =   12
      Top             =   105
      Width           =   4770
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name:"
      Height          =   240
      Left            =   135
      TabIndex        =   11
      Top             =   480
      Width           =   1365
   End
   Begin VB.Label Label24 
      BackStyle       =   0  'Transparent
      Caption         =   "Middle Name:"
      Height          =   240
      Left            =   135
      TabIndex        =   10
      Top             =   855
      Width           =   1425
   End
   Begin VB.Label Label28 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name:"
      Height          =   240
      Left            =   135
      TabIndex        =   9
      Top             =   1200
      Width           =   1470
   End
End
Attribute VB_Name = "frmAddPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mlPersonID As Long, mbAddMode As Boolean

Event PersonUpdated(ThePerson As Long)

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim rs As Recordset
On Error GoTo ErrorTrap
    
    If PersonalFieldsValidatedOK Then
    
        Set rs = CMSDB.OpenRecordset("tblNameAddress", dbOpenDynaset)
        
        With rs
        
        If mlPersonID = 0 Then
            'new rec
            .AddNew
        Else
            .FindFirst "ID = " & mlPersonID
            .Edit
        End If
        
        !FirstName = txtFirstName
        !OfficialFirstName = txtFirstName
        !LastName = txtLastName
        !MiddleName = txtMiddleName
        !GenderMF = "M"
        !DOB = 0
        ![InfirmityLevel0-6] = 0
        !Active = True
        !Anointed = False
        !LinkedAddressPerson = 0
        
        !HomePhone = Trim(txtHomePhone)
        
        If Trim(txtMobile) = "" And Trim(txtMobile2) <> "" Then
            txtMobile = txtMobile2
            txtMobile2 = ""
        End If
        
        !MobilePhone = Trim(txtMobile)
        !MobilePhone2 = Trim(txtMobile2)
        !Email = Trim(txtEmail)
        
        .Update
        .Requery
            
        End With
        
        Set rs = CMSDB.OpenRecordset("SELECT MAX(ID) AS MaxID FROM tblNameAddress", dbOpenDynaset)
        
        If mlPersonID = 0 Then
            mlPersonID = rs!MaxID
        End If
        
        rs.Close
        Set rs = Nothing
        
        RaiseEvent PersonUpdated(mlPersonID)
        
        If FormIsOpen("frmPersonalDetails") Then
            With frmPersonalDetails
            
            If .chkActiveOnly = vbUnchecked Then
                .RefreshNamesList
            End If
            
            End With
        End If
    
        '
        'Update the cache of names
        '
        CongregationMember.LoadNamesToCache
        
        Unload Me
        
    End If

    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub

Private Function PersonalFieldsValidatedOK() As Boolean
        
Dim rstNameCheck As Recordset, Query As String
Dim TheFirstName As String, TheLastName As String, TheMiddleName As String


On Error GoTo ErrorTrap

    PersonalFieldsValidatedOK = True
        
    'Validate txtFirstName
        
    If Len(txtFirstName) = 0 Then
        PersonalFieldsValidatedOK = False
        MsgBox "First name should be between 1 and 50 characters long. " & _
                "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel + vbExclamation, AppName
            
        txtFirstName.SetFocus
        Exit Function
    Else
        PersonalFieldsValidatedOK = True
    End If
    
    'Validate txtLastName

    If Len(txtLastName) = 0 Then
        PersonalFieldsValidatedOK = False
        MsgBox "Last name should be between 1 and 50 characters long. " & _
                  "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel + vbExclamation, AppName
            
        txtLastName.SetFocus
        Exit Function
    Else
        PersonalFieldsValidatedOK = True
    End If
    
           
    'Check if person's name already appears in tblNameAddress - only if it's changed
    
    TheFirstName = DoubleUpSingleQuotes(txtFirstName)
    TheMiddleName = DoubleUpSingleQuotes(txtMiddleName)
    TheLastName = DoubleUpSingleQuotes(txtLastName)
    
    'this IF is done simply because, if txtMiddleName is blank the SQL wouldn't work
    If txtMiddleName = "" Or IsNull(txtMiddleName) Then
        Query = "SELECT * FROM tblNameAddress WHERE FirstName = '" & TheFirstName & "'" & " And LastName = '" & TheLastName & "'" & " AND (IsNull(MiddleName) OR Len(MiddleName)) = 0"
    Else
        Query = "SELECT * FROM tblNameAddress WHERE FirstName = '" & TheFirstName & "'" & " And LastName = '" & TheLastName & "'" & " AND MiddleName = '" & TheMiddleName & "'"
    End If
    
    If mlPersonID > 0 Then
        Query = Query & " AND ID <> " & mlPersonID
    End If

    Set rstNameCheck = CMSDB.OpenRecordset(Query, dbOpenDynaset)

    If Not rstNameCheck.BOF Then
        PersonalFieldsValidatedOK = False
        MsgBox txtFirstName & " " & txtMiddleName & " " & txtLastName & " already exists." & _
                  " Try to make it unique by altering the Middle Name, for example." & _
                  " Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel + vbExclamation, AppName
        txtMiddleName.SetFocus
        rstNameCheck.Close
        Set rstNameCheck = Nothing
        Exit Function
    Else
        rstNameCheck.Close
        PersonalFieldsValidatedOK = True
        Set rstNameCheck = Nothing
    End If
            
    Exit Function
ErrorTrap:
    EndProgram
    
End Function


Private Sub Form_Load()

On Error GoTo ErrorTrap
    
    If mlPersonID > 0 Then
        'update mode
        With CongregationMember
        
        txtFirstName = .GetFirstName(mlPersonID)
        txtMiddleName = .GetMiddleName(mlPersonID)
        txtLastName = .GetLastName(mlPersonID)
        txtHomePhone = .HomePhone(mlPersonID)
        txtMobile = .MobilePhone(mlPersonID)
        txtMobile2 = .MobilePhone2(mlPersonID)
        txtEmail = .Email(mlPersonID)
        
        End With
    Else
        txtFirstName = ""
        txtMiddleName = ""
        txtLastName = ""
    End If
    
    Exit Sub
ErrorTrap:
    Call EndProgram
End Sub


Public Property Get PersonID() As Long
    PersonID = mlPersonID
End Property

Public Property Let PersonID(ByVal vNewValue As Long)
    mlPersonID = vNewValue
End Property
Public Property Get AddMode() As Boolean
    AddMode = mbAddMode
End Property

Public Property Let AddMode(ByVal vNewValue As Boolean)
    mbAddMode = vNewValue
End Property
Private Sub txtEmail_GotFocus()
    TextFieldGotFocus txtEmail
End Sub

Private Sub txtFirstName_GotFocus()
    TextFieldGotFocus txtFirstName
End Sub
Private Sub txtHomePhone_GotFocus()
    TextFieldGotFocus txtHomePhone
End Sub

Private Sub txtLastName_GotFocus()
    TextFieldGotFocus txtLastName
End Sub
Private Sub txtMiddleName_GotFocus()
    TextFieldGotFocus txtMiddleName
End Sub

Private Sub txtMobile_GotFocus()
    TextFieldGotFocus txtMobile
End Sub
Private Sub txtMobile2_GotFocus()
    TextFieldGotFocus txtMobile2
End Sub

