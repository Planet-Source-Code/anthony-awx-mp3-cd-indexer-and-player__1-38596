VERSION 5.00
Begin VB.Form frmConfig 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Configuration"
   ClientHeight    =   7110
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   9120
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   3600
      ScaleHeight     =   285
      ScaleWidth      =   840
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   5820
      Width           =   870
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   15
         ScaleHeight     =   240
         ScaleWidth      =   795
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   15
         Width           =   795
         Begin VB.ComboBox cmbPlayerDrive 
            Height          =   315
            ItemData        =   "frmConfig.frx":0442
            Left            =   -30
            List            =   "frmConfig.frx":045E
            TabIndex        =   29
            Text            =   "d:\"
            Top             =   -30
            Width           =   870
         End
      End
   End
   Begin VB.CheckBox chkAutoSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Open Search Window at Startup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5010
      TabIndex        =   23
      Top             =   4485
      Width           =   3120
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Submit"
      Height          =   375
      Left            =   6705
      TabIndex        =   20
      Top             =   2160
      Width           =   960
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   0
      Left            =   3600
      ScaleHeight     =   285
      ScaleWidth      =   840
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4530
      Width           =   870
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   15
         ScaleHeight     =   240
         ScaleWidth      =   795
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   15
         Width           =   795
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmConfig.frx":0489
            Left            =   -30
            List            =   "frmConfig.frx":049C
            TabIndex        =   19
            Text            =   "1"
            Top             =   -30
            Width           =   870
         End
      End
   End
   Begin VB.OptionButton optInterface 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Multiple Window"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   4260
      TabIndex        =   11
      Top             =   3615
      Width           =   2370
   End
   Begin VB.OptionButton optInterface 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Single Window"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   945
      TabIndex        =   10
      Top             =   3615
      Value           =   -1  'True
      Width           =   2370
   End
   Begin VB.TextBox txtDBFile 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   705
      TabIndex        =   7
      Top             =   2175
      Width           =   5910
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Check this box if you want a search window to appear automatically each time you start Disk Pro."
      ForeColor       =   &H00404040&
      Height          =   705
      Index           =   6
      Left            =   5280
      TabIndex        =   30
      Top             =   4830
      Width           =   2670
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "What Drive do you want the MP3 Player to look for songs?"
      ForeColor       =   &H00404040&
      Height          =   480
      Index           =   5
      Left            =   690
      TabIndex        =   26
      Top             =   6150
      Width           =   3000
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Onboard MP3 Player"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   3
      Left            =   690
      TabIndex        =   25
      Top             =   5865
      Width           =   1710
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808080&
      Height          =   15
      Index           =   2
      Left            =   690
      TabIndex        =   24
      Top             =   5760
      Width           =   7260
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808080&
      Height          =   15
      Index           =   1
      Left            =   705
      TabIndex        =   22
      Top             =   4440
      Width           =   7260
   End
   Begin VB.Label Label8 
      BackColor       =   &H00808080&
      Height          =   15
      Index           =   0
      Left            =   705
      TabIndex        =   21
      Top             =   3000
      Width           =   7260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search Window Default"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   2
      Left            =   705
      TabIndex        =   16
      Top             =   4545
      Width           =   1950
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "How many ""Search"" windows do you want to open each time you click the ""Search"" Icon on the toolbar?"
      ForeColor       =   &H00404040&
      Height          =   705
      Index           =   4
      Left            =   705
      TabIndex        =   15
      Top             =   4830
      Width           =   3045
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Each Window is opened tiled above the previous window."
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   3
      Left            =   4530
      TabIndex        =   14
      Top             =   3885
      Width           =   3105
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Each window is maximized by default."
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   2
      Left            =   1215
      TabIndex        =   13
      Top             =   3885
      Width           =   2700
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This setting allows you to select the default interface used when you open windows in Disk Pro."
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   1
      Left            =   705
      TabIndex        =   12
      Top             =   3300
      Width           =   6855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Interface Preference"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   1
      Left            =   705
      TabIndex        =   9
      Top             =   3105
      Width           =   1770
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is the name of the file that contains the data for the CD's you will be indexing."
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   0
      Left            =   705
      TabIndex        =   8
      Top             =   2550
      Width           =   5940
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Database Filename"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   705
      TabIndex        =   6
      Top             =   1935
      Width           =   1620
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Use the settings below to set up Disk Pro to work the way you do"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   765
      TabIndex        =   5
      Top             =   630
      Width           =   4695
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   45
      Left            =   0
      TabIndex        =   2
      Top             =   1650
      Width           =   1.00005e5
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Configuration"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   705
      TabIndex        =   1
      Top             =   165
      Width           =   1755
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   15
      Left            =   705
      TabIndex        =   0
      Top             =   540
      Width           =   1.00005e5
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmConfig.frx":04AF
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Configuration"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   180
      Width           =   1755
   End
   Begin VB.Image imgLogo 
      Height          =   1650
      Left            =   5520
      Picture         =   "frmConfig.frx":08F1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3570
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   1.00005e5
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAutoSearch_Click()

    Config.iSearchAtStartup = chkAutoSearch.Value
    SaveConfig
    
End Sub

Private Sub cmbPlayerDrive_Change()

    Config.iDefaultPlayerDrive = Trim$(cmbPlayerDrive.Text)
    SaveConfig
    
End Sub

Private Sub cmbPlayerDrive_Click()

    cmbPlayerDrive_Change
    
End Sub

Private Sub cmbPlayerDrive_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Private Sub cmdSubmit_Click()

    Config.DBFileName = Trim$(Me.txtDBFile.Text)
    SaveConfig

End Sub

Private Sub Combo1_Change()

    Config.iSearchWindows = Combo1.Text
    SaveConfig
    frmMDI.lblTool(0).Caption = "Search (" & Config.iSearchWindows & ")"

    
End Sub

Private Sub Combo1_Click()

    Combo1_Change
    
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Private Sub Form_Load()

    txtDBFile = Trim$(Config.DBFileName)
    Me.optInterface(Config.iInterface).Value = True
    If Config.iSearchWindows > 0 Then Me.Combo1.Text = Config.iSearchWindows
    If Config.iSearchAtStartup = 1 Then Me.chkAutoSearch.Value = Checked
    Me.cmbPlayerDrive.Text = Trim$(Config.iDefaultPlayerDrive)
    
End Sub

Private Sub Form_Resize()

    With imgLogo
        .Top = 15
        .Left = ScaleWidth - .Width - 15
    End With
    
End Sub

Private Sub optInterface_Click(Index As Integer)

    Config.iInterface = Index
    SaveConfig
    
End Sub


Private Sub txtDBFile_Change()

    cmdSubmit.Default = True
    
End Sub
