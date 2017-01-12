VERSION 5.00
Begin VB.Form frmSuspendDatesMaint 
   Caption         =   "C.M.S. View / Amend Suspend Dates"
   ClientHeight    =   8805
   ClientLeft      =   345
   ClientTop       =   0
   ClientWidth     =   14325
   ScaleHeight     =   8805
   ScaleWidth      =   14325
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkDelFlag00 
      Height          =   223
      Left            =   330
      TabIndex        =   160
      Top             =   2625
      Width           =   164
   End
   Begin VB.CheckBox chkDelFlag01 
      Height          =   223
      Left            =   332
      TabIndex        =   161
      Top             =   2910
      Width           =   164
   End
   Begin VB.CheckBox chkDelFlag02 
      Height          =   223
      Left            =   332
      TabIndex        =   162
      Top             =   3195
      Width           =   164
   End
   Begin VB.CheckBox chkDelFlag03 
      Height          =   223
      Left            =   332
      TabIndex        =   163
      Top             =   3480
      Width           =   164
   End
   Begin VB.CheckBox chkDelFlag04 
      Height          =   223
      Left            =   332
      TabIndex        =   164
      Top             =   3750
      Width           =   164
   End
   Begin VB.CheckBox chkDelFlag05 
      Height          =   223
      Left            =   332
      TabIndex        =   165
      Top             =   4035
      Width           =   164
   End
   Begin VB.CheckBox chkDelFlag06 
      Height          =   223
      Left            =   332
      TabIndex        =   166
      Top             =   4320
      Width           =   164
   End
   Begin VB.CheckBox chkDelFlag07 
      Height          =   223
      Left            =   332
      TabIndex        =   167
      Top             =   4605
      Width           =   164
   End
   Begin VB.CheckBox chkDelFlag08 
      Height          =   223
      Left            =   332
      TabIndex        =   168
      Top             =   4905
      Width           =   164
   End
   Begin VB.CheckBox chkDelFlag09 
      Height          =   223
      Left            =   332
      TabIndex        =   169
      Top             =   5190
      Width           =   164
   End
   Begin VB.CheckBox chkDelFlag10 
      Height          =   223
      Left            =   332
      TabIndex        =   170
      Top             =   5445
      Width           =   164
   End
   Begin VB.CheckBox chkDelFlag11 
      Height          =   223
      Left            =   332
      TabIndex        =   171
      Top             =   5730
      Width           =   164
   End
   Begin VB.CheckBox chkDelFlag12 
      Height          =   223
      Left            =   332
      TabIndex        =   172
      Top             =   6045
      Width           =   164
   End
   Begin VB.CheckBox chkDelFlag13 
      Height          =   223
      Left            =   332
      TabIndex        =   173
      Top             =   6330
      Width           =   164
   End
   Begin VB.CheckBox chkDelFlag14 
      Height          =   223
      Left            =   332
      TabIndex        =   174
      Top             =   6600
      Width           =   164
   End
   Begin VB.TextBox txtTaskCat10 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   61
      Top             =   5385
      Width           =   2268
   End
   Begin VB.TextBox txtTaskCat11 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   67
      Top             =   5669
      Width           =   2268
   End
   Begin VB.TextBox txtTaskSubCat07 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   4535
      Width           =   2267
   End
   Begin VB.TextBox txtTaskSubCat08 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   4818
      Width           =   2267
   End
   Begin VB.TextBox txtTaskSubCat11 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   68
      Top             =   5669
      Width           =   2267
   End
   Begin VB.TextBox txtTask11 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   69
      Top             =   5669
      Width           =   2268
   End
   Begin VB.TextBox txtTask04 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   3685
      Width           =   2268
   End
   Begin VB.TextBox txtTask05 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   33
      Top             =   3968
      Width           =   2268
   End
   Begin VB.TextBox txtReason07 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   4535
      Width           =   2551
   End
   Begin VB.TextBox txtReason09 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   60
      Top             =   5102
      Width           =   2551
   End
   Begin VB.TextBox txtTaskCat01 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2834
      Width           =   2268
   End
   Begin VB.TextBox txtTaskCat02 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   3118
      Width           =   2268
   End
   Begin VB.TextBox txtTaskCat03 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   3401
      Width           =   2268
   End
   Begin VB.TextBox txtTaskCat04 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   25
      Top             =   3685
      Width           =   2268
   End
   Begin VB.TextBox txtTaskCat05 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   3968
      Width           =   2268
   End
   Begin VB.TextBox txtTaskCat06 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   37
      Top             =   4251
      Width           =   2268
   End
   Begin VB.TextBox txtTaskCat07 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   4535
      Width           =   2268
   End
   Begin VB.TextBox txtTaskCat08 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   4818
      Width           =   2268
   End
   Begin VB.TextBox txtTaskCat09 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   5102
      Width           =   2268
   End
   Begin VB.TextBox txtTaskCat12 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   73
      Top             =   5952
      Width           =   2268
   End
   Begin VB.TextBox txtTaskCat13 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   79
      Top             =   6236
      Width           =   2268
   End
   Begin VB.TextBox txtTaskCat14 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   85
      Top             =   6519
      Width           =   2268
   End
   Begin VB.TextBox txtTask07 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   45
      Top             =   4535
      Width           =   2268
   End
   Begin VB.TextBox txtTask08 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   4818
      Width           =   2268
   End
   Begin VB.TextBox txtTaskSubCat09 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   5102
      Width           =   2267
   End
   Begin VB.TextBox txtTask09 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   5102
      Width           =   2268
   End
   Begin VB.TextBox txtTaskSubCat10 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   62
      Top             =   5385
      Width           =   2267
   End
   Begin VB.TextBox txtTask10 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   63
      Top             =   5385
      Width           =   2268
   End
   Begin VB.TextBox txtTaskSubCat12 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   74
      Top             =   5952
      Width           =   2267
   End
   Begin VB.TextBox txtTask12 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   75
      Top             =   5952
      Width           =   2268
   End
   Begin VB.TextBox txtTaskSubCat13 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   80
      Top             =   6236
      Width           =   2267
   End
   Begin VB.TextBox txtTask13 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   81
      Top             =   6519
      Width           =   2268
   End
   Begin VB.TextBox txtTaskSubCat14 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   86
      Top             =   6519
      Width           =   2267
   End
   Begin VB.TextBox txtTask14 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   87
      Top             =   6236
      Width           =   2268
   End
   Begin VB.TextBox txtTaskSubCat01 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2834
      Width           =   2267
   End
   Begin VB.TextBox txtTask01 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   2834
      Width           =   2268
   End
   Begin VB.TextBox txtTaskSubCat02 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   3118
      Width           =   2267
   End
   Begin VB.TextBox txtTask02 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   3118
      Width           =   2268
   End
   Begin VB.TextBox txtTaskSubCat03 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   3401
      Width           =   2267
   End
   Begin VB.TextBox txtTask03 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   21
      Top             =   3401
      Width           =   2268
   End
   Begin VB.TextBox txtTaskSubCat04 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   3685
      Width           =   2267
   End
   Begin VB.TextBox txtTaskSubCat05 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   32
      Top             =   3968
      Width           =   2267
   End
   Begin VB.TextBox txtTaskSubCat06 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   4251
      Width           =   2267
   End
   Begin VB.TextBox txtTask06 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   39
      Top             =   4251
      Width           =   2268
   End
   Begin VB.TextBox txtReason01 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2834
      Width           =   2551
   End
   Begin VB.TextBox txtReason02 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   3118
      Width           =   2551
   End
   Begin VB.TextBox txtReason03 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   3401
      Width           =   2551
   End
   Begin VB.TextBox txtReason04 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   3685
      Width           =   2551
   End
   Begin VB.TextBox txtReason05 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   36
      Top             =   3968
      Width           =   2551
   End
   Begin VB.TextBox txtReason06 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   42
      Top             =   4251
      Width           =   2551
   End
   Begin VB.TextBox txtReason08 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   4818
      Width           =   2551
   End
   Begin VB.TextBox txtReason10 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   5385
      Width           =   2551
   End
   Begin VB.TextBox txtReason11 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   72
      Top             =   5669
      Width           =   2551
   End
   Begin VB.TextBox txtReason12 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   78
      Top             =   5952
      Width           =   2551
   End
   Begin VB.TextBox txtReason13 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   84
      Top             =   6236
      Width           =   2551
   End
   Begin VB.TextBox txtReason14 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   90
      Top             =   6519
      Width           =   2551
   End
   Begin VB.TextBox txtTaskSubCat00 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2551
      Width           =   2267
   End
   Begin VB.TextBox txtReason00 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2551
      Width           =   2551
   End
   Begin VB.TextBox txtTask00 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   2551
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskCat07 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   122
      Top             =   4535
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskCat08 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   126
      Top             =   4818
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskCat09 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   130
      Top             =   5102
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskCat10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   134
      Top             =   5385
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskCat11 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   138
      Top             =   5669
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskCat12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   142
      Top             =   5952
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskCat13 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   146
      Top             =   6236
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskCat14 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   150
      Top             =   6519
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskCat01 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   98
      Top             =   2834
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskCat02 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   102
      Top             =   3118
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskCat03 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   106
      Top             =   3401
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskCat04 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   110
      Top             =   3685
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskCat05 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   114
      Top             =   3968
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskCat06 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   118
      Top             =   4251
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.TextBox txtTaskCat00 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2551
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskCat00 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   566
      Locked          =   -1  'True
      TabIndex        =   95
      Top             =   2551
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskSubCat07 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   123
      Top             =   4535
      Visible         =   0   'False
      Width           =   2267
   End
   Begin VB.ComboBox cmbTask07 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   124
      Top             =   4535
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskSubCat08 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   127
      Top             =   4818
      Visible         =   0   'False
      Width           =   2267
   End
   Begin VB.ComboBox cmbTask08 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   128
      Top             =   4818
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskSubCat09 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   131
      Top             =   5102
      Visible         =   0   'False
      Width           =   2267
   End
   Begin VB.ComboBox cmbTask09 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   132
      Top             =   5102
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskSubCat10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   135
      Top             =   5385
      Visible         =   0   'False
      Width           =   2267
   End
   Begin VB.ComboBox cmbTask10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   136
      Top             =   5385
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskSubCat11 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   139
      Top             =   5669
      Visible         =   0   'False
      Width           =   2267
   End
   Begin VB.ComboBox cmbTask11 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   140
      Top             =   5669
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskSubCat12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   143
      Top             =   5952
      Visible         =   0   'False
      Width           =   2267
   End
   Begin VB.ComboBox cmbTask12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   144
      Top             =   5952
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskSubCat13 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   147
      Top             =   6236
      Visible         =   0   'False
      Width           =   2267
   End
   Begin VB.ComboBox cmbTask13 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   148
      Top             =   6236
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbTaskSubCat14 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   151
      Top             =   6519
      Visible         =   0   'False
      Width           =   2267
   End
   Begin VB.ComboBox cmbTask14 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   152
      Top             =   6519
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbReason07 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   125
      Top             =   4535
      Visible         =   0   'False
      Width           =   2551
   End
   Begin VB.ComboBox cmbReason08 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   129
      Top             =   4818
      Visible         =   0   'False
      Width           =   2551
   End
   Begin VB.ComboBox cmbReason09 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   133
      Top             =   5102
      Visible         =   0   'False
      Width           =   2551
   End
   Begin VB.ComboBox cmbReason10 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   137
      Top             =   5385
      Visible         =   0   'False
      Width           =   2551
   End
   Begin VB.ComboBox cmbReason11 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   141
      Top             =   5669
      Visible         =   0   'False
      Width           =   2551
   End
   Begin VB.ComboBox cmbReason12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   145
      Top             =   5952
      Visible         =   0   'False
      Width           =   2551
   End
   Begin VB.ComboBox cmbReason13 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   149
      Top             =   6236
      Visible         =   0   'False
      Width           =   2551
   End
   Begin VB.ComboBox cmbReason14 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   153
      Top             =   6519
      Visible         =   0   'False
      Width           =   2551
   End
   Begin VB.ComboBox cmbTaskSubCat01 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   99
      Top             =   2834
      Visible         =   0   'False
      Width           =   2267
   End
   Begin VB.ComboBox cmbTask01 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   100
      Top             =   2834
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbReason01 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   101
      Top             =   2834
      Visible         =   0   'False
      Width           =   2551
   End
   Begin VB.ComboBox cmbTaskSubCat02 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   103
      Top             =   3118
      Visible         =   0   'False
      Width           =   2267
   End
   Begin VB.ComboBox cmbTask02 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   104
      Top             =   3118
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbReason02 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   105
      Top             =   3118
      Visible         =   0   'False
      Width           =   2551
   End
   Begin VB.ComboBox cmbTaskSubCat03 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   107
      Top             =   3401
      Visible         =   0   'False
      Width           =   2267
   End
   Begin VB.ComboBox cmbTask03 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   108
      Top             =   3401
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbReason03 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   109
      Top             =   3401
      Visible         =   0   'False
      Width           =   2551
   End
   Begin VB.ComboBox cmbTaskSubCat04 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   111
      Top             =   3685
      Visible         =   0   'False
      Width           =   2267
   End
   Begin VB.ComboBox cmbTask04 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   112
      Top             =   3685
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbReason04 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   113
      Top             =   3685
      Visible         =   0   'False
      Width           =   2551
   End
   Begin VB.ComboBox cmbTaskSubCat05 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   115
      Top             =   3968
      Visible         =   0   'False
      Width           =   2267
   End
   Begin VB.ComboBox cmbTask05 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   116
      Top             =   3968
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbReason05 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   117
      Top             =   3968
      Visible         =   0   'False
      Width           =   2551
   End
   Begin VB.ComboBox cmbTaskSubCat06 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   119
      Top             =   4251
      Visible         =   0   'False
      Width           =   2267
   End
   Begin VB.ComboBox cmbTask06 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   120
      Top             =   4251
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbReason06 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   121
      Top             =   4251
      Visible         =   0   'False
      Width           =   2551
   End
   Begin VB.TextBox txtStartDate01 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   7369
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   2834
      Width           =   1134
   End
   Begin VB.TextBox txtEndDate01 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   8503
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2834
      Width           =   1134
   End
   Begin VB.TextBox txtStartDate02 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   7369
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   3118
      Width           =   1134
   End
   Begin VB.TextBox txtEndDate02 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   8503
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   3118
      Width           =   1134
   End
   Begin VB.TextBox txtStartDate03 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   7369
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   3401
      Width           =   1134
   End
   Begin VB.TextBox txtEndDate03 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   8503
      Locked          =   -1  'True
      TabIndex        =   23
      Top             =   3401
      Width           =   1134
   End
   Begin VB.TextBox txtStartDate04 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   7369
      Locked          =   -1  'True
      TabIndex        =   28
      Top             =   3685
      Width           =   1134
   End
   Begin VB.TextBox txtEndDate04 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   8503
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   3685
      Width           =   1134
   End
   Begin VB.TextBox txtStartDate05 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   7369
      Locked          =   -1  'True
      TabIndex        =   34
      Top             =   3968
      Width           =   1134
   End
   Begin VB.TextBox txtEndDate05 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   8503
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   3968
      Width           =   1134
   End
   Begin VB.TextBox txtStartDate06 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   7369
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   4251
      Width           =   1134
   End
   Begin VB.TextBox txtEndDate06 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   8503
      Locked          =   -1  'True
      TabIndex        =   41
      Top             =   4251
      Width           =   1134
   End
   Begin VB.TextBox txtStartDate07 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   7369
      Locked          =   -1  'True
      TabIndex        =   46
      Top             =   4535
      Width           =   1134
   End
   Begin VB.TextBox txtEndDate07 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   8503
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   4535
      Width           =   1134
   End
   Begin VB.TextBox txtStartDate08 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   7369
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   4818
      Width           =   1134
   End
   Begin VB.TextBox txtEndDate08 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   8503
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   4818
      Width           =   1134
   End
   Begin VB.TextBox txtStartDate09 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   7369
      Locked          =   -1  'True
      TabIndex        =   58
      Top             =   5102
      Width           =   1134
   End
   Begin VB.TextBox txtEndDate09 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   8503
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   5102
      Width           =   1134
   End
   Begin VB.TextBox txtStartDate10 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   7369
      Locked          =   -1  'True
      TabIndex        =   64
      Top             =   5385
      Width           =   1134
   End
   Begin VB.TextBox txtEndDate10 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   8503
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   5385
      Width           =   1134
   End
   Begin VB.TextBox txtStartDate11 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   7369
      Locked          =   -1  'True
      TabIndex        =   70
      Top             =   5669
      Width           =   1134
   End
   Begin VB.TextBox txtEndDate11 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   8503
      Locked          =   -1  'True
      TabIndex        =   71
      Top             =   5669
      Width           =   1134
   End
   Begin VB.TextBox txtStartDate12 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   7369
      Locked          =   -1  'True
      TabIndex        =   76
      Top             =   5952
      Width           =   1134
   End
   Begin VB.TextBox txtEndDate12 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   8503
      Locked          =   -1  'True
      TabIndex        =   77
      Top             =   5952
      Width           =   1134
   End
   Begin VB.TextBox txtStartDate13 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   7369
      Locked          =   -1  'True
      TabIndex        =   82
      Top             =   6236
      Width           =   1134
   End
   Begin VB.TextBox txtEndDate13 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   8503
      Locked          =   -1  'True
      TabIndex        =   83
      Top             =   6236
      Width           =   1134
   End
   Begin VB.TextBox txtStartDate14 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   7369
      Locked          =   -1  'True
      TabIndex        =   88
      Top             =   6519
      Width           =   1134
   End
   Begin VB.TextBox txtEndDate14 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   314
      Left            =   8503
      Locked          =   -1  'True
      TabIndex        =   89
      Top             =   6519
      Width           =   1134
   End
   Begin VB.ComboBox cmbTaskSubCat00 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   2834
      Locked          =   -1  'True
      TabIndex        =   96
      Top             =   2551
      Visible         =   0   'False
      Width           =   2267
   End
   Begin VB.ComboBox cmbTask00 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   5101
      Locked          =   -1  'True
      TabIndex        =   97
      Top             =   2551
      Visible         =   0   'False
      Width           =   2268
   End
   Begin VB.ComboBox cmbReason00 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   283
      Left            =   9637
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2551
      Visible         =   0   'False
      Width           =   2551
   End
   Begin VB.TextBox txtStartDate00 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   7369
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2551
      Width           =   1134
   End
   Begin VB.TextBox txtEndDate00 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   313
      Left            =   8503
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2551
      Width           =   1134
   End
   Begin VB.CommandButton cmdCloseThisThing 
      Caption         =   "Cancel"
      Height          =   570
      Left            =   12755
      TabIndex        =   91
      Top             =   1984
      Width           =   1125
   End
   Begin VB.ComboBox cmbCongregation 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   566
      TabIndex        =   92
      Top             =   566
      Width           =   3375
   End
   Begin VB.ComboBox cmbPerson 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   284
      Left            =   566
      TabIndex        =   93
      Top             =   1416
      Width           =   3972
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   570
      Left            =   12755
      TabIndex        =   154
      Top             =   283
      Width           =   1125
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   570
      Left            =   12755
      TabIndex        =   155
      Top             =   4251
      Width           =   1125
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Edit"
      Height          =   570
      Left            =   12755
      TabIndex        =   156
      Top             =   5102
      Width           =   1125
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   570
      Left            =   12755
      TabIndex        =   157
      Top             =   5952
      Width           =   1125
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "Browse"
      Enabled         =   0   'False
      Height          =   570
      Left            =   12755
      TabIndex        =   158
      Top             =   3401
      Width           =   1125
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Enabled         =   0   'False
      Height          =   570
      Left            =   12755
      TabIndex        =   159
      Top             =   1133
      Width           =   1125
   End
   Begin VB.CheckBox chkApplyToAll 
      Caption         =   "Apply to all"
      Enabled         =   0   'False
      Height          =   240
      Left            =   5640
      TabIndex        =   175
      Top             =   780
      Width           =   225
   End
   Begin VB.CheckBox chkDeleteAll 
      Caption         =   "Delete all"
      Height          =   240
      Left            =   5640
      TabIndex        =   176
      Top             =   1185
      Width           =   225
   End
   Begin VB.Shape box00 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   283
      Left            =   240
      Top             =   2550
      Width           =   330
   End
   Begin VB.Shape Box02 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   283
      Left            =   240
      Top             =   3120
      Width           =   330
   End
   Begin VB.Shape Box03 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   283
      Left            =   240
      Top             =   3405
      Width           =   330
   End
   Begin VB.Shape Box04 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   283
      Left            =   240
      Top             =   3690
      Width           =   330
   End
   Begin VB.Shape Box05 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   268
      Left            =   240
      Top             =   3975
      Width           =   330
   End
   Begin VB.Shape Box06 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   283
      Left            =   240
      Top             =   4245
      Width           =   330
   End
   Begin VB.Shape Box07 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   283
      Left            =   240
      Top             =   4530
      Width           =   330
   End
   Begin VB.Shape Box08 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   283
      Left            =   240
      Top             =   4815
      Width           =   330
   End
   Begin VB.Shape Box09 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   283
      Left            =   240
      Top             =   5100
      Width           =   330
   End
   Begin VB.Shape Box10 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   283
      Left            =   240
      Top             =   5385
      Width           =   330
   End
   Begin VB.Shape Box11 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   283
      Left            =   240
      Top             =   5670
      Width           =   330
   End
   Begin VB.Shape Box12 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   283
      Left            =   240
      Top             =   5955
      Width           =   330
   End
   Begin VB.Shape Box13 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   283
      Left            =   240
      Top             =   6240
      Width           =   330
   End
   Begin VB.Shape Box14 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   283
      Left            =   240
      Top             =   6525
      Width           =   330
   End
   Begin VB.Shape Box01 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      Height          =   283
      Left            =   240
      Top             =   2835
      Width           =   330
   End
   Begin VB.Label Label139 
      BackStyle       =   0  'Transparent
      Caption         =   "Task Category"
      ForeColor       =   &H00000000&
      Height          =   302
      Left            =   1133
      TabIndex        =   94
      Top             =   2267
      Width           =   1105
   End
   Begin VB.Label Label140 
      BackStyle       =   0  'Transparent
      Caption         =   "Task Subcategory"
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   3401
      TabIndex        =   177
      Top             =   2267
      Width           =   1380
   End
   Begin VB.Label Label141 
      BackStyle       =   0  'Transparent
      Caption         =   "Task"
      ForeColor       =   &H00000000&
      Height          =   302
      Left            =   5952
      TabIndex        =   178
      Top             =   2267
      Width           =   415
   End
   Begin VB.Label Label142 
      BackStyle       =   0  'Transparent
      Caption         =   "Start Date"
      ForeColor       =   &H00000000&
      Height          =   302
      Left            =   7369
      TabIndex        =   179
      Top             =   2267
      Width           =   805
   End
   Begin VB.Label Label143 
      BackStyle       =   0  'Transparent
      Caption         =   "End Date"
      ForeColor       =   &H00000000&
      Height          =   302
      Left            =   8503
      TabIndex        =   180
      Top             =   2267
      Width           =   805
   End
   Begin VB.Label Label144 
      BackStyle       =   0  'Transparent
      Caption         =   "Suspend Reason"
      ForeColor       =   &H00000000&
      Height          =   302
      Left            =   9920
      TabIndex        =   181
      Top             =   2267
      Width           =   1795
   End
   Begin VB.Label Label88 
      BackStyle       =   0  'Transparent
      Caption         =   "Del"
      ForeColor       =   &H00000000&
      Height          =   302
      Left            =   255
      TabIndex        =   182
      Top             =   2265
      Width           =   340
   End
   Begin VB.Label Label150 
      BackStyle       =   0  'Transparent
      Caption         =   "Congregation:"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   566
      TabIndex        =   183
      Top             =   283
      Width           =   1065
   End
   Begin VB.Label Label152 
      BackStyle       =   0  'Transparent
      Caption         =   "Person:"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   566
      TabIndex        =   184
      Top             =   1133
      Width           =   1065
   End
End
Attribute VB_Name = "frmSuspendDatesMaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rstTaskDates As Recordset, DisableScrollBarOnUpdateEvent As Boolean, InEditMode As Boolean
Dim TaskIndexArray(14) As TaskIXDef



Private Sub ActiveXCtl1_Updated(Code As Integer)
    If Not DisableScrollBarOnUpdateEvent Then   'This check is done since form load seems to fire a ScrollBarUpdated event, causing
                                'FillInTheFields to be executed again!
        Call FillInTheFields(0, 0, True, False, False)
    End If
    
    DisableScrollBarOnUpdateEvent = False

End Sub

Private Sub chkDeleteAll_AfterUpdate()

    If chkDeleteAll.Value = True Then
        TickDelFlags
    Else
        ClearDelFlags
    End If

End Sub


Private Sub cmbCongregation_AfterUpdate()
    Me!cmbPerson = Null
    Call FillInTheFields(0, 0, False, False, True)
    Me!cmbCongregation.SetFocus
End Sub

Private Sub cmbCongregation_Change()
    Me!cmbPerson.RowSource = "SELECT tblNameAddress.ID, " & _
                         "tblNameAddress.FirstName & ' ' &" & _
                         "tblNameAddress.MiddleName & ' ' &" & _
                         "tblNameAddress.LastName " & _
                  "FROM tblNameAddress " & _
                  "WHERE ID IN (SELECT Person FROM tblTaskAndPerson WHERE CongNo = " & Me!cmbCongregation & ")"
                      
    Me!cmbPerson = Null
    Call FillInTheFields(0, 0, False, True, True)
    
End Sub

Private Sub cmbCongregation_NotInList(NewData As String, Response As Integer)
    Me!cmbCongregation.Undo
    Response = acDataErrContinue
End Sub

Private Sub cmbPerson_AfterUpdate()
    Call FillInTheFields(0, 0, False, True, False)
End Sub

Private Sub cmbPerson_Change()
    Set rstTaskDates = GenerateTaskDateRecordset(Me!cmbPerson, Me!cmbCongregation)
    Call FillInTheFields(0, 0, False, True, False)
End Sub

Private Sub cmbPerson_NotInList(NewData As String, Response As Integer)
    Me!cmbPerson.Undo
    Response = acDataErrContinue
End Sub

Private Sub cmbReason00_AfterUpdate()
    Me!cmdApply.Enabled = True
End Sub

Private Sub cmbReason00_NotInList(NewData As String, Response As Integer)
    Me!cmbReason00.Undo
    Response = acDataErrContinue
End Sub

Private Sub cmbTask00_AfterUpdate()
    Me!cmdApply.Enabled = True
End Sub

Private Sub cmbTask00_NotInList(NewData As String, Response As Integer)
    Me!cmbTask00.Undo
    Response = acDataErrContinue

End Sub

Private Sub cmbTaskCat00_AfterUpdate()
    Me!cmbTaskSubCat00.RowSource = "SELECT TaskSubCategory, Description FROM tblTaskSubCategories WHERE CongNo = " & Me!cmbCongregation & " AND TaskCategory = " & Me!cmbTaskCat00
    Me!cmbTaskSubCat00.Requery
    Me!cmbTaskSubCat00 = Null
    Me!cmdApply.Enabled = True
End Sub

Private Sub cmbTaskCat00_NotInList(NewData As String, Response As Integer)
    Me!cmbTaskCat00.Undo
    Response = acDataErrContinue
End Sub

Private Sub cmbTaskSubCat00_AfterUpdate()
    Me!cmbTask00.RowSource = "SELECT Task, Description FROM tblTasks " & _
                            " WHERE CongNo = " & Me!cmbCongregation & _
                            " AND TaskCategory = " & Me!cmbTaskCat00 & _
                            " AND TaskSubCategory = " & Me!cmbTaskSubCat00 & _
                            " AND Task IN (SELECT Task FROM tblTaskAndPerson " & _
                                                "WHERE Person = " & Me!cmbPerson & _
                                                " AND CongNo = " & Me!cmbCongregation & _
                                                " AND TaskCategory = " & Me!cmbTaskCat00 & _
                                                " AND TaskSubCategory = " & Me!cmbTaskSubCat00 & ")" & _
                            " ORDER BY Description"
                            
    Me!cmbTask00.Requery
    
    Me!cmdApply.Enabled = True
                            

End Sub


Private Sub cmbTaskSubCat00_NotInList(NewData As String, Response As Integer)
    Me!cmbTaskSubCat00.Undo
    Response = acDataErrContinue
End Sub

Private Sub cmdAdd_Click()
Dim i As Integer, iStr As String
      
    DisableScrollBarOnUpdateEvent = True
          
    'Disable lower 14 rows, change top row to combos
    For i = 1 To 14 Step 1
            
        iStr = Format(i, "00")
        
        Me.Controls("txtTask" & iStr).Locked = False
        Me.Controls("txtTaskCat" & iStr).Locked = False
        Me.Controls("txtTaskSubCat" & iStr).Locked = False
        Me.Controls("txtReason" & iStr).Locked = False
        Me.Controls("txtStartDate" & iStr).Locked = False
        Me.Controls("txtEndDate" & iStr).Locked = False
        
        Me.Controls("txtTask" & iStr).Enabled = False
        Me.Controls("txtTaskCat" & iStr).Enabled = False
        Me.Controls("txtTaskSubCat" & iStr).Enabled = False
        Me.Controls("txtReason" & iStr).Enabled = False
        Me.Controls("txtStartDate" & iStr).Enabled = False
        Me.Controls("txtEndDate" & iStr).Enabled = False

    Next i

    Me.Controls("txtTask00").Visible = False
    Me.Controls("txtTaskCat00").Visible = False
    Me.Controls("txtTaskSubCat00").Visible = False
    Me.Controls("txtReason00").Visible = False
    
    Me.Controls("txtStartDate00").Visible = True
    Me.Controls("txtEndDate00").Visible = True
    Me.Controls("txtStartDate00").Enabled = True
    Me.Controls("txtEndDate00").Enabled = True
    
    Me.Controls("txtTask00").Locked = False
    Me.Controls("txtTaskCat00").Locked = False
    Me.Controls("txtTaskSubCat00").Locked = False
    Me.Controls("txtReason00").Locked = False
    Me.Controls("txtStartDate00").Locked = False
    Me.Controls("txtEndDate00").Locked = False
    
    Me.Controls("cmbTask00").Visible = True
    Me.Controls("cmbTaskCat00").Visible = True
    Me.Controls("cmbTaskSubCat00").Visible = True
    Me.Controls("cmbReason00").Visible = True
    
    Me.Controls("cmbTask00").Locked = False
    Me.Controls("cmbTaskCat00").Locked = False
    Me.Controls("cmbTaskSubCat00").Locked = False
    Me.Controls("cmbReason00").Locked = False

    
    'enable/disable other controls etc
    Me!cmbCongregation.Enabled = False
    Me!cmbPerson.Enabled = False
    Me!cmdBrowse.Enabled = True
    Me!cmdOK.Enabled = False
    Me!cmbTaskCat00.SetFocus
    Me!cmdAdd.Enabled = False
    Me!cmdEdit.Enabled = False
    Me!cmdDelete.Enabled = False
    Me!chkApplyToAll.Enabled = True
    Me!chkDeleteAll.Enabled = False
    
    InEditMode = True
    
    Set rstTaskDates = GenerateTaskDateRecordset(cmbPerson, cmbCongregation)
    'fill fields in grid
    Call FillInTheFields(1, 0, False, False, False)

    PopulateCombos
            
    'disable checkboxes and their background

    For i = 0 To 14 Step 1
        'Convert i to a 2-character text string with leading 0
        iStr = Format(i, "00")
        
        Me.Controls("chkDelFlag" & iStr).Enabled = False
        Me.Controls("box" & iStr).BackColor = 12632256
    Next i
    
            
End Sub

Private Sub cmdApply_Click()
Dim ShallWeSaveAndExit As Boolean
    
    DisableScrollBarOnUpdateEvent = True
    
    Call ApplyChanges(ShallWeSaveAndExit)

End Sub

Private Sub cmdBrowse_Click()
    DisableScrollBarOnUpdateEvent = True
    ToBrowseMode
    Set rstTaskDates = GenerateTaskDateRecordset(cmbPerson, cmbCongregation)
    Call FillInTheFields(0, 0, False, True, False)
End Sub

Private Sub cmdCloseThisThing_Click()
    'MsgBox "Contains: " & Forms!frmsuspenddatesmaint.Controls(0), vbOKOnly, AppName
    DoCmd.Close acForm, "frmSuspendDatesMaint"
End Sub

Private Sub cmdDelete_Click()
    
    DisableScrollBarOnUpdateEvent = True
    
    DoDeletes
    
    Set rstTaskDates = GenerateTaskDateRecordset(cmbPerson, cmbCongregation)
    
    Call FillInTheFields(0, 0, False, True, False)
    
    DisableScrollBarOnUpdateEvent = True
    
    'move scrollbar to top
    Me!ActiveXCtl1.Value = 0
    
    If chkDeleteAll.Value = True Then
        TickDelFlags
    Else
        ClearDelFlags
    End If
     
End Sub

Private Sub cmdOK_Click()
    Dim CheckIfWeSaveAndExit As Boolean
      
    If InEditMode Then
        Call ApplyChanges(CheckIfWeSaveAndExit)
        If CheckIfWeSaveAndExit Then
            DoCmd.Close acForm, "frmSuspendDatesMaint"
        End If
    Else
        DoCmd.Close acForm, "frmSuspendDatesMaint"
    End If
        
        
End Sub

Private Sub Form_Activate()
    'DisableScrollBarOnUpdateEvent = False
End Sub

Private Sub Form_Load()
Dim i As Integer, iStr As String, TheControl As Control
    
    DisableScrollBarOnUpdateEvent = True

'Initialise display of various text/combo boxes
    ToBrowseMode

'Populate the main combos and select entry corresponding to frmPersonalDetails

    Me!cmbCongregation.RowSource = "SELECT CongNo, CongName FROM tblCong"
    Me!cmbCongregation = frmPersonalDetails.cmbCongregation
    Me!cmbCongregation.SetFocus
    
    Me!cmbPerson.RowSource = "SELECT tblNameAddress.ID, " & _
                         "tblNameAddress.FirstName & ' ' &" & _
                         "tblNameAddress.MiddleName & ' ' &" & _
                         "tblNameAddress.LastName " & _
                        "FROM tblNameAddress " & _
                        "WHERE ID IN (SELECT Person FROM tblTaskAndPerson WHERE CongNo = " & Me!cmbCongregation & ")"
                                 
    Me!cmbPerson = frmPersonalDetails.lstNames
    
    Set rstTaskDates = GenerateTaskDateRecordset([frmPersonalDetails].[lstNames], [frmPersonalDetails].[cmbCongregation])
    
    ToBrowseMode
    
    Call FillInTheFields(0, 0, False, True, False)
    
    ClearDelFlags
            
End Sub

Private Sub FillInTheFields(PopulateFromLine As Integer, StartFromRecordNumber As Integer, IsScrollEvent As Boolean, CheckIfScrollBarRequired As Boolean, SetAllFieldsBlank As Boolean)
Dim APieceOfSQL As String, i As Integer, j As Integer, RecNumber As Integer
Dim iStr As String
    
       
    With rstTaskDates
    'Get recordcount
    If Not .BOF Then
        .MoveLast
        .MoveFirst
    End If
    
    'Check if scroll-bar is required
    
    
    If CheckIfScrollBarRequired Then
        If .RecordCount + PopulateFromLine <= 15 Then
            Me!ActiveXCtl1.Enabled = False
        Else
            Me!ActiveXCtl1.Enabled = True
        End If
    Else
        If Not IsScrollEvent Then
            Me!ActiveXCtl1.Enabled = False
        End If
    End If
            
    
    'Has this routine been called by Scroll-bar event?
    If IsScrollEvent Then
        'Check position of scroll bar, use this to find the first record to display...
        RecNumber = Int((Me!ActiveXCtl1.Value / Me!ActiveXCtl1.Max) * .RecordCount)
    Else
        RecNumber = StartFromRecordNumber
    End If
    
    If Not .BOF Then
    
        .MoveFirst
        
        If .RecordCount + PopulateFromLine > 15 Then
            If RecNumber > .RecordCount Then
                .MoveLast
            Else
                .Move (RecNumber)
            End If
        End If
        
    End If
        
    'Now fill all text fields, starting from row specified by PopulateFromLine
    For i = 0 To 14 Step 1
        
        'Convert i to a 2-character text string with leading 0
        iStr = Format(i, "00")
        
        If i < PopulateFromLine Or SetAllFieldsBlank Or .EOF Or .BOF Then
            Me.Controls("txtTaskCat" & iStr) = ""
            Me.Controls("txtTaskSubCat" & iStr) = ""
            Me.Controls("txtTask" & iStr) = ""
            Me.Controls("txtStartDate" & iStr) = ""
            Me.Controls("txtEndDate" & iStr) = ""
            Me.Controls("txtReason" & iStr) = ""
            Me.Controls("chkDelFlag" & iStr).Value = False
            Me.Controls("chkDelFlag" & iStr).Visible = False
            
            'Move zero key to array to keep track of where we are on grid
            TaskIndexArray(i).SeqNum = 0
            TaskIndexArray(i).TaskCat = 0
            TaskIndexArray(i).TaskSubCat = 0
            TaskIndexArray(i).Task = 0
            
        Else
            'assign fields from rstTaskDate to txt fields
            Me.Controls("txtTaskCat" & iStr) = .Fields(0)
            Me.Controls("txtTaskSubCat" & iStr) = .Fields(1)
            Me.Controls("txtTask" & iStr) = .Fields(2)
            Me.Controls("txtStartDate" & iStr) = .Fields(3)
            Me.Controls("txtEndDate" & iStr) = .Fields(4)
            Me.Controls("txtReason" & iStr) = .Fields(5)
            Me.Controls("chkDelFlag" & iStr).Visible = True
            If chkDeleteAll.Value = True Then
                Me.Controls("chkDelFlag" & iStr).Value = True
            Else
                Me.Controls("chkDelFlag" & iStr).Value = False
            End If
            
            'Move key to array to keep track of where we are on grid
            TaskIndexArray(i).SeqNum = .Fields(6)
            TaskIndexArray(i).TaskCat = .Fields(7)
            TaskIndexArray(i).TaskSubCat = .Fields(8)
            TaskIndexArray(i).Task = .Fields(9)

            .MoveNext
            
        End If

            
            'If .EOF Then
            '    Exit For
            'End If
            
    Next i
    
    ClearDelFlags
    
    End With
                      
End Sub

Private Sub ToBrowseMode()
Dim i As Integer, iStr As String
    
    'Make overlaid combos invisible
    'Make overlaid text-boxes visible
    For i = 0 To 14 Step 1
            
        iStr = Format(i, "00")
        Me.Controls("cmbTask" & iStr).Visible = False
        Me.Controls("cmbTaskCat" & iStr).Visible = False
        Me.Controls("cmbTaskSubCat" & iStr).Visible = False
        Me.Controls("cmbReason" & iStr).Visible = False
        
        Me.Controls("txtTask" & iStr).Visible = True
        Me.Controls("txtTaskCat" & iStr).Visible = True
        Me.Controls("txtTaskSubCat" & iStr).Visible = True
        Me.Controls("txtReason" & iStr).Visible = True
        Me.Controls("txtStartDate" & iStr).Visible = True
        Me.Controls("txtEndDate" & iStr).Visible = True

        Me.Controls("txtTask" & iStr).Enabled = False
        Me.Controls("txtTaskCat" & iStr).Enabled = False
        Me.Controls("txtTaskSubCat" & iStr).Enabled = False
        Me.Controls("txtReason" & iStr).Enabled = False
        Me.Controls("txtStartDate" & iStr).Enabled = False
        Me.Controls("txtEndDate" & iStr).Enabled = False

        
        Me.Controls("txtTask" & iStr).Locked = True
        Me.Controls("txtTaskCat" & iStr).Locked = True
        Me.Controls("txtTaskSubCat" & iStr).Locked = True
        Me.Controls("txtReason" & iStr).Locked = True
        Me.Controls("txtStartDate" & iStr).Locked = True
        Me.Controls("txtEndDate" & iStr).Locked = True
        
    Next i
    
    Me!cmbPerson.Enabled = True
    Me!cmbCongregation.Enabled = True
    Me!cmdOK.Enabled = True
    Me!cmdAdd.Enabled = True
    Me!cmdAdd.SetFocus
    Me!cmdBrowse.Enabled = False
    Me!cmdEdit.Enabled = True
    Me!cmdDelete.Enabled = True
    Me!cmdApply.Enabled = False
    Me!chkApplyToAll.Value = False
    Me!chkApplyToAll.Enabled = False
    Me!chkDeleteAll.Value = False
    Me!chkDeleteAll.Enabled = True

    'enable checkboxes and their background

    For i = 0 To 14 Step 1
        'Convert i to a 2-character text string with leading 0
        iStr = Format(i, "00")
        
        Me.Controls("chkDelFlag" & iStr).Enabled = True
        Me.Controls("box" & iStr).BackColor = vbWhite
    Next i
    
    ClearDelFlags

    InEditMode = False
            
End Sub

Private Function GenerateTaskDateRecordset(Person As Long, Congregation As Long) As Recordset
Dim APieceOfSQL As String
       
    APieceOfSQL = "SELECT tblTaskCategories.Description, " & _
                 "tblTaskSubCategories.Description, " & _
                 "tblTasks.Description, " & _
                 "tblTaskPersonSuspendDates.SuspendStartDate, " & _
                 "tblTaskPersonSuspendDates.SuspendEndDate, " & _
                 "tblSuspendReasons.SuspendReasonDesc, " & _
                 "tblTaskPersonSuspendDates.SeqNo, " & _
                 "tblTaskPersonSuspendDates.TaskCategory, " & _
                 "tblTaskPersonSuspendDates.TaskSubCategory, " & _
                 "tblTaskPersonSuspendDates.Task " & _
                 "FROM (((tblTaskCategories INNER JOIN tblTaskSubCategories ON " & _
                 "(tblTaskCategories.TaskCategory = tblTaskSubCategories.TaskCategory) AND " & _
                 "(tblTaskCategories.CongNo = tblTaskSubCategories.CongNo)) " & _
                 "INNER JOIN tblTasks ON (tblTaskSubCategories.TaskSubCategory = tblTasks.TaskSubCategory) " & _
                 "AND (tblTaskSubCategories.TaskCategory = tblTasks.TaskCategory) " & _
                 "AND (tblTaskSubCategories.CongNo = tblTasks.CongNo)) " & _
                 "INNER JOIN tblTaskPersonSuspendDates ON (tblTasks.TaskSubCategory = tblTaskPersonSuspendDates.TaskSubCategory) " & _
                 "AND (tblTasks.TaskCategory = tblTaskPersonSuspendDates.TaskCategory) " & _
                 "AND (tblTasks.CongNo = tblTaskPersonSuspendDates.CongNo) " & _
                 "AND (tblTasks.Task = tblTaskPersonSuspendDates.Task)) " & _
                 "INNER JOIN tblSuspendReasons ON tblTaskPersonSuspendDates.SuspendReason = tblSuspendReasons.SuspendReasonCode " & _
                 "WHERE tblTaskPersonSuspendDates.CongNo = " & Congregation & _
                " AND tblTaskPersonSuspendDates.Person = " & Person & _
                " ORDER By SuspendStartDate DESC;"
     
                  
    Set GenerateTaskDateRecordset = CMSDB.OpenRecordset(APieceOfSQL, dbOpenDynaset)
    

End Function

Private Sub PopulateCombos()
    Me!cmbTaskCat00.RowSource = "SELECT TaskCategory, Description FROM tblTaskCategories WHERE CongNo = " & Me!cmbCongregation
    Me!cmbTaskCat00 = Null
    Me!cmbReason00.RowSource = "SELECT SuspendReasonCode, SuspendReasonDesc FROM tblSuspendReasons"
    Me!cmbTaskSubCat00.RowSource = ""
    Me!cmbTask00.RowSource = ""
    Me!cmbTask00 = Null
    Me!cmbReason00 = Null
    
End Sub

Private Sub txtEndDate00_AfterUpdate()
    Me!cmdApply.Enabled = True
End Sub

Private Sub txtEndDate00_Exit(Cancel As Integer)
    Me!txtEndDate00 = Trim(Me!txtEndDate00)

    If IsDate(Me!txtEndDate00) Then
        Me!txtEndDate00 = Format(Me!txtEndDate00, "Short Date")
    End If
    
    Cancel = False


End Sub

Private Sub txtStartDate00_AfterUpdate()
    Me!cmdApply.Enabled = True
End Sub

Private Function EntryValidatedOK()
     'Validate txtStartDate00

    If Not IsDate(Me!txtStartDate00) Then
        EntryValidatedOK = False
        If MsgBox("Date is not valid. " & _
                  "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
        Else
            'field is invalid, so forget entire changes and go to browse-mode since user clicked Cancel.
            ToBrowseMode
        End If
        Me!txtStartDate00.SetFocus
        Exit Function
    Else
        EntryValidatedOK = True
    End If
    
    'Validate txtEndDate00

    If Not IsDate(Me!txtEndDate00) Then
        EntryValidatedOK = False
        If MsgBox("Date is not valid. " & _
                  "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
        Else
            'field is invalid, so forget entire changes and go to browse-mode since user clicked Cancel.
            ToBrowseMode
        End If
        Me!txtEndDate00.SetFocus
        Exit Function
    Else
        EntryValidatedOK = True
    End If
    

    'Is enddate < startdate?

    If CDate(Me!txtEndDate00) < CDate(Me!txtStartDate00) Then
        EntryValidatedOK = False
        If MsgBox("The End-date is earlier than the Start-date! " & _
                  "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
        Else
            'field is invalid, so forget entire changes and go to browse-mode since user clicked Cancel.
            ToBrowseMode
        End If
        Me!txtStartDate00.SetFocus
        Exit Function
    Else
        EntryValidatedOK = True
    End If

    'Are combos selected?

    If IsNull(Me!cmbTask00) Or IsNull(Me!cmbTaskCat00) Or IsNull(Me!cmbTaskSubCat00) Or IsNull(Me!cmbReason00) Then
        EntryValidatedOK = False
        If MsgBox("You must select Role Category, Subcategory, Role and Suspend Reason. " & _
                  "Click OK to correct field, Cancel to abandon all changes to this record.", vbOKCancel, AppName) = vbOK Then
        Else
            'field is invalid, so forget entire changes and go to browse-mode since user clicked Cancel.
            ToBrowseMode
        End If
        Me!cmbTaskCat00.SetFocus
        Exit Function
    Else
        EntryValidatedOK = True
    End If


End Function

Private Sub txtStartDate00_Exit(Cancel As Integer)
    Me!txtStartDate00 = Trim(Me!txtStartDate00)

    If IsDate(Me!txtStartDate00) Then
        Me!txtStartDate00 = Format(Me!txtStartDate00, "Short Date")
    End If
    
    Cancel = False

End Sub

Private Sub ApplyChanges(SaveAndExit As Boolean)
Dim StartDateUS As String, EndDateUS As String, rstTaskPers As Recordset, TaskCount As Integer, i As Integer

    If EntryValidatedOK Then
    
        StartDateUS = Format(Me!txtStartDate00, "mm/dd/yy")
        EndDateUS = Format(Me!txtEndDate00, "mm/dd/yy")
        
        If Me!chkApplyToAll.Value = True Then
            'Update for all of Person's tasks
            
            
            Set rstTaskPers = CMSDB.OpenRecordset("SELECT * FROM tblTaskAndPerson " & _
                                                        " WHERE Person = " & Me!cmbPerson & _
                                                        " AND CongNo = " & Me!cmbCongregation _
                                                        , dbOpenDynaset)
            
            With rstTaskPers
            .MoveFirst
            .MoveLast
            .MoveFirst
            TaskCount = .RecordCount
            
            For i = 1 To TaskCount
                DeleteTaskWithBlankDates !TaskCategory, !TaskSubCategory, !Task
                
                CMSDB.Execute "INSERT INTO tblTaskPersonSuspendDates " & _
                              "(CongNo, TaskCategory, TaskSubCategory, Task, Person, SuspendStartDate, SuspendEndDate, SuspendReason) " & _
                              "VALUES (" & Me!cmbCongregation & ", " & _
                                     !TaskCategory & ", " & _
                                     !TaskSubCategory & ", " & _
                                     !Task & ", " & _
                                     Me!cmbPerson & ", #" & _
                                     StartDateUS & "#, #" & _
                                     EndDateUS & "#, " & _
                                     Me!cmbReason00 & ")"
                                     
                .MoveNext
            Next i
            
            .Close
            
            End With
            
        Else
            
            DeleteTaskWithBlankDates Me!cmbTaskCat00, Me!cmbTaskSubCat00, Me!cmbTask00
            
            CMSDB.Execute "INSERT INTO tblTaskPersonSuspendDates " & _
                          "(CongNo, TaskCategory, TaskSubCategory, Task, Person, SuspendStartDate, SuspendEndDate, SuspendReason) " & _
                          "VALUES (" & Me!cmbCongregation & ", " & _
                                 Me!cmbTaskCat00 & ", " & _
                                 Me!cmbTaskSubCat00 & ", " & _
                                 Me!cmbTask00 & ", " & _
                                 Me!cmbPerson & ", #" & _
                                 StartDateUS & "#, #" & _
                                 EndDateUS & "#, " & _
                                 Me!cmbReason00 & ")"
        End If
        
        
        'fill fields in grid with blanks
        Call FillInTheFields(0, 0, False, False, True)
        
        DisableScrollBarOnUpdateEvent = True 'stop the damn scroll-bar from misbehaving!
        
        'Populate the fields
        Set rstTaskDates = GenerateTaskDateRecordset(cmbPerson, cmbCongregation)
        Call FillInTheFields(1, 0, False, True, False)
            
        PopulateCombos
                             
        Me!cmbTaskCat00.SetFocus
        Me!cmdApply.Enabled = False
        
        SaveAndExit = True
    Else
        SaveAndExit = False
    End If
        

End Sub

Private Sub ClearDelFlags()
Dim i As Integer, iStr As String

    For i = 0 To 14 Step 1
        'Convert i to a 2-character text string with leading 0
        iStr = Format(i, "00")
        
        Me.Controls("chkDelFlag" & iStr).Value = False
        
    Next i
        
End Sub

Private Sub DoDeletes()
Dim i As Integer, iStr As String, rstTempRSet As Recordset

    'scan all chkBoxes and for checked ones find corresponding entry on TaskIndexArray. Use this to do delete from tblTaskPersonSuspendDates
    For i = 0 To 14 Step 1
        'Convert i to a 2-character text string with leading 0
        iStr = Format(i, "00")
        
        
        
        If Me.Controls("chkDelFlag" & iStr).Value = True Then
            CMSDB.Execute ("DELETE FROM tblTaskPersonSuspendDates " & _
                                "WHERE SeqNo = " & TaskIndexArray(i).SeqNum & _
                                " AND CongNo = " & Me!cmbCongregation & _
                                " AND TaskCategory = " & TaskIndexArray(i).TaskCat & _
                                " AND TaskSubCategory = " & TaskIndexArray(i).TaskSubCat & _
                                " AND Task = " & TaskIndexArray(i).Task & _
                                " AND Person = " & Me!cmbPerson)
                                    
            Set rstTempRSet = CMSDB.OpenRecordset("SELECT * FROM tblTaskPersonSuspendDates " & _
                                                        " WHERE Person = " & Me!cmbPerson & _
                                                        " AND CongNo = " & Me!cmbCongregation & _
                                                        " AND TaskCategory = " & TaskIndexArray(i).TaskCat & _
                                                        " AND TaskSubCategory = " & TaskIndexArray(i).TaskSubCat & _
                                                        " AND Task = " & TaskIndexArray(i).Task & _
                                                        " AND Person = " & Me!cmbPerson, dbOpenDynaset)
                
            
            '
            'If task no longer exists on tblTaskPersonSuspendDates, insert it with blank suspend dates
            '
            If rstTempRSet.BOF Then
                CMSDB.Execute "INSERT INTO tblTaskPersonSuspendDates " & _
                              "(CongNo, TaskCategory, TaskSubCategory, Task, Person) " & _
                              "VALUES (" & Me!cmbCongregation & ", " & _
                                     TaskIndexArray(i).TaskCat & ", " & _
                                     TaskIndexArray(i).TaskSubCat & ", " & _
                                     TaskIndexArray(i).Task & ", " & _
                                     Me!cmbPerson & ")"
            End If
                                
            rstTempRSet.Close
        End If
    Next i
  
End Sub

Private Sub TickDelFlags()
Dim i As Integer, iStr As String

    For i = 0 To 14 Step 1
        'Convert i to a 2-character text string with leading 0
        iStr = Format(i, "00")
        If Not IsNull(Me.Controls("txtTaskCat" & iStr)) And Me.Controls("txtTaskCat" & iStr) <> "" Then
            Me.Controls("chkDelFlag" & iStr).Value = True
        End If
    Next i

End Sub

Private Function TaskHasBlankSuspendDates(TaskCat As Integer, TaskSubCat As Integer, Task As Integer) As Boolean
Dim rstTempRecSet As Recordset

    Set rstTempRecSet = CMSDB.OpenRecordset("SELECT * FROM tblTaskPersonSuspendDates " & _
                                                " WHERE Person = " & Me!cmbPerson & _
                                                " AND CongNo = " & Me!cmbCongregation & _
                                                " AND TaskCategory = " & TaskCat & _
                                                " AND TaskSubCategory = " & TaskSubCat & _
                                                " AND Task = " & Task & _
                                                " AND Person = " & Me!cmbPerson & _
                                                " AND IsNull(SuspendStartDate)" & _
                                                " AND IsNull(SuspendEndDate)", dbOpenDynaset)
                                                
    If rstTempRecSet.BOF Then
        TaskHasBlankSuspendDates = False
    Else
        TaskHasBlankSuspendDates = True
    End If

End Function




Private Sub DeleteTaskWithBlankDates(TaskCat As Integer, TaskSubCat As Integer, Task As Integer)

    CMSDB.Execute "DELETE FROM tblTaskPersonSuspendDates " & _
                                        " WHERE Person = " & Me!cmbPerson & _
                                        " AND CongNo = " & Me!cmbCongregation & _
                                        " AND TaskCategory = " & TaskCat & _
                                        " AND TaskSubCategory = " & TaskSubCat & _
                                        " AND Task = " & Task & _
                                        " AND Person = " & Me!cmbPerson & _
                                        " AND IsNull(SuspendStartDate)" & _
                                        " AND IsNull(SuspendEndDate)"

End Sub


