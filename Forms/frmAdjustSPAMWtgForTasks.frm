VERSION 5.00
Begin VB.Form frmAdjustSPAMWtgForTasks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C.M.S. Adjust SPAM Weightings for Individual Tasks"
   ClientHeight    =   4065
   ClientLeft      =   330
   ClientTop       =   -15
   ClientWidth     =   6630
   Icon            =   "frmAdjustSPAMWtgForTasks.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSPAMWeighting 
      Alignment       =   2  'Center
      ForeColor       =   &H00000000&
      Height          =   329
      Left            =   1908
      MaxLength       =   4
      TabIndex        =   3
      Top             =   3480
      Width           =   850
   End
   Begin VB.ListBox lstTaskCategory 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   840
      Left            =   150
      TabIndex        =   0
      Top             =   495
      Width           =   2608
   End
   Begin VB.ListBox lstTaskSubCategory 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1425
      Left            =   150
      TabIndex        =   1
      Top             =   1845
      Width           =   2608
   End
   Begin VB.ListBox lstTask 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2790
      Left            =   2970
      TabIndex        =   2
      Top             =   495
      Width           =   3448
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   453
      Left            =   2970
      TabIndex        =   4
      Top             =   3420
      Width           =   1134
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   453
      Left            =   4125
      TabIndex        =   5
      Top             =   3420
      Width           =   1134
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   453
      Left            =   5280
      TabIndex        =   6
      Top             =   3420
      Width           =   1134
   End
   Begin VB.Label Label1 
      Caption         =   "SPAM Weighting"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   10
      Top             =   3525
      Width           =   1485
   End
   Begin VB.Label lblCat 
      BackStyle       =   0  'Transparent
      Caption         =   "Role Category:"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   150
      TabIndex        =   7
      Top             =   150
      Width           =   1635
   End
   Begin VB.Label lblSubCat 
      BackStyle       =   0  'Transparent
      Caption         =   "Role Sub Category:"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   180
      TabIndex        =   8
      Top             =   1545
      Width           =   1875
   End
   Begin VB.Label lblTask 
      BackStyle       =   0  'Transparent
      Caption         =   "Role:"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2970
      TabIndex        =   9
      Top             =   150
      Width           =   1065
   End
End
Attribute VB_Name = "frmAdjustSPAMWtgForTasks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ChangeMade As Boolean, FormLoading As Boolean


Private Sub cmdApply_Click()
Dim SaveAndExit As Boolean

On Error GoTo ErrorTrap

    Call ApplyChanges(SaveAndExit)
    

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


Private Sub cmdOK_Click()
    Dim SaveAndExit As Boolean
    

On Error GoTo ErrorTrap

    Call ApplyChanges(SaveAndExit)
    
    If SaveAndExit Then
        Unload Me
    End If
    
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub

Private Sub Form_Load()


On Error GoTo ErrorTrap

    FormLoading = True
    
    lstTask.ListIndex = -1
        
    lblCat.FontBold = True
    lblSubCat.FontBold = False
    lblTask.FontBold = False

    lblCat.ForeColor = 255
    lblSubCat.ForeColor = 0
    lblTask.ForeColor = 0

    HandleListBox.PopulateListBox Me!lstTaskCategory, "SELECT TaskCategory, Description FROM tblTaskCategories", _
                                  CMSDB, 0, "", True, 1
    lstTaskSubCategory.Clear
    lstTask.Clear
    
    
    lblCat.FontBold = True
    lblSubCat.FontBold = False
    lblTask.FontBold = False
    
    lblCat.ForeColor = 255
    lblSubCat.ForeColor = 0
    lblTask.ForeColor = 0
    
    txtSPAMWeighting = ""
    
    FormLoading = False

    Exit Sub
ErrorTrap:
    EndProgram
End Sub


Private Sub Form_Unload(Cancel As Integer)
    BringForwardMainMenuWhenItsTheLastFormOpen
End Sub



Private Sub lstTask_Click()
Dim SQLString As String, rstTemp As Recordset

On Error GoTo ErrorTrap


    SQLString = "SELECT SPAMRotaWeighting FROM tblTaskWeightings " & _
                "WHERE TaskCategory = " & lstTaskCategory.ItemData(lstTaskCategory.ListIndex) & _
               " AND TaskSubCategory = " & lstTaskSubCategory.ItemData(lstTaskSubCategory.ListIndex) & _
               " AND Task = " & lstTask.ItemData(lstTask.ListIndex)
               
    Set rstTemp = CMSDB.OpenRecordset(SQLString, dbOpenSnapshot)
       
    FormLoading = True
    
    If Not rstTemp.BOF Then
        txtSPAMWeighting = rstTemp!SPAMRotaWeighting
    Else
        txtSPAMWeighting = 0
    End If
    
    rstTemp.Close
    
    FormLoading = False
    

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub lstTaskCategory_Click()

On Error GoTo ErrorTrap

    HandleListBox.PopulateListBox Me!lstTaskSubCategory, _
                                "SELECT TaskSubCategory, Description " & _
                                "FROM tblTaskSubCategories " & _
                                "WHERE TaskCategory = " & lstTaskCategory.ItemData(lstTaskCategory.ListIndex), _
                                    CMSDB, 0, "", True, 1
    lstTask.Clear

    lstTaskSubCategory.SetFocus
    lblCat.FontBold = False
    lblSubCat.FontBold = True
    lblTask.FontBold = False
    
    lblCat.ForeColor = 0
    lblSubCat.ForeColor = 255
    lblTask.ForeColor = 0

    txtSPAMWeighting = ""
    
    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub


Private Sub lstTaskSubCategory_click()
Dim SQLDingDong As String

On Error GoTo ErrorTrap


    SQLDingDong = "SELECT Task, Description FROM tblTasks " & _
                "WHERE TaskCategory = " & lstTaskCategory.ItemData(lstTaskCategory.ListIndex) & _
               " AND TaskSubCategory = " & lstTaskSubCategory.ItemData(lstTaskSubCategory.ListIndex)

    HandleListBox.PopulateListBox Me!lstTask, SQLDingDong, CMSDB, 0, "", True, 1
    lstTask.SetFocus
    lblCat.FontBold = False
    lblSubCat.FontBold = False
    lblTask.FontBold = True

    lblCat.ForeColor = 0
    lblSubCat.ForeColor = 0
    lblTask.ForeColor = 255
    
    txtSPAMWeighting = ""
    

    Exit Sub
ErrorTrap:
    EndProgram
    
End Sub



Private Function ApplyChanges(SaveAndExit As Boolean)

On Error GoTo ErrorTrap

    txtSPAMWeighting = Trim(txtSPAMWeighting)
    
    If ChangeMade Then
        If (Len(txtSPAMWeighting) = 0 Or IsNull(txtSPAMWeighting)) Then
            MsgBox "Weighting field should not be blank", vbOKOnly + vbInformation, AppName
            txtSPAMWeighting.SetFocus
            SaveAndExit = False
        ElseIf Not IsNumeric(txtSPAMWeighting) Then
            MsgBox "Weighting field should be numeric", vbOKOnly + vbInformation, AppName
            txtSPAMWeighting.SetFocus
            SaveAndExit = False
        Else
            If MsgBox("Are you sure you want to save this change?", vbYesNo + vbQuestion, _
                        AppName) = vbYes Then
                
                UpdateTask
                
                txtSPAMWeighting.text = ""
                
                ChangeMade = False
                lstTaskCategory.SetFocus
                cmdApply.Enabled = False
                cmdOK.Enabled = False
                
                SaveAndExit = True
            Else
                SaveAndExit = False
            End If
        End If
    Else
        SaveAndExit = True
    End If


    Exit Function
ErrorTrap:
    EndProgram
    
End Function




Private Sub UpdateTask()
Dim UpdateSQL As String


On Error GoTo ErrorTrap

    
    UpdateSQL = "UPDATE tblTaskWeightings " & _
                "SET SPAMRotaWeighting = " & CInt(txtSPAMWeighting) & _
               " WHERE TaskCategory = " & lstTaskCategory.ItemData(lstTaskCategory.ListIndex) & _
               " AND TaskSubCategory = " & lstTaskSubCategory.ItemData(lstTaskSubCategory.ListIndex) & _
               " AND Task = " & lstTask.ItemData(lstTask.ListIndex)

    CMSDB.Execute UpdateSQL
    
    HandleListBox.Requery Me!lstTask, False
    

    Exit Sub
ErrorTrap:
    EndProgram
    

End Sub


Private Sub txtSPAMWeighting_Change()
    If FormLoading Then Exit Sub
    
    cmdApply.Enabled = True
    cmdOK.Enabled = True
    ChangeMade = True
End Sub

Private Sub txtSPAMWeighting_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorTrap

'Must be numeric. Allow Backspace (8)  Delete and arrow keys seem to be allowed by default.

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    Exit Sub
ErrorTrap:
    EndProgram


End Sub
