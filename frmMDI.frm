VERSION 5.00
Begin VB.MDIForm frmMDI 
   Appearance      =   0  'Flat
   BackColor       =   &H8000000C&
   Caption         =   "Disk Pro Plus"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   10980
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H00B49800&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7800
      Left            =   0
      ScaleHeight     =   520
      ScaleMode       =   0  'User
      ScaleWidth      =   45.327
      TabIndex        =   0
      Top             =   0
      Width           =   1710
      Begin VB.CommandButton cmdAlign 
         Caption         =   "&>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   1410
         TabIndex        =   5
         Top             =   75
         Width           =   225
      End
      Begin VB.CommandButton cmdAlign 
         Caption         =   "&<"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   1200
         TabIndex        =   6
         Top             =   75
         Width           =   225
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Height          =   2.45745e5
         Index           =   1
         Left            =   -360
         TabIndex        =   10
         Top             =   0
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Height          =   2.45745e5
         Index           =   0
         Left            =   1695
         TabIndex        =   9
         Top             =   -495
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Toolbar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   90
         Width           =   645
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Height          =   345
         Left            =   15
         TabIndex        =   7
         Top             =   30
         Width           =   1635
      End
      Begin VB.Label lblTool 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Configure"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   435
         MouseIcon       =   "frmMDI.frx":0442
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   3495
         Width           =   825
      End
      Begin VB.Image imgTool 
         Height          =   480
         Index           =   2
         Left            =   607
         MouseIcon       =   "frmMDI.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "frmMDI.frx":0A56
         Top             =   2940
         Width           =   480
      End
      Begin VB.Label lblTool 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add a CD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   465
         MouseIcon       =   "frmMDI.frx":0E98
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   2280
         Width           =   765
      End
      Begin VB.Label lblTool 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   540
         MouseIcon       =   "frmMDI.frx":11A2
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   1035
         Width           =   615
      End
      Begin VB.Image imgTool 
         Height          =   480
         Index           =   1
         Left            =   607
         MouseIcon       =   "frmMDI.frx":14AC
         MousePointer    =   99  'Custom
         Picture         =   "frmMDI.frx":17B6
         Top             =   1725
         Width           =   480
      End
      Begin VB.Image imgTool 
         Height          =   480
         Index           =   0
         Left            =   607
         MouseIcon       =   "frmMDI.frx":1BF8
         MousePointer    =   99  'Custom
         Picture         =   "frmMDI.frx":1F02
         Top             =   570
         Width           =   480
      End
      Begin VB.Label lblHighlight 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   195
         MouseIcon       =   "frmMDI.frx":2344
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   405
         Visible         =   0   'False
         Width           =   1290
      End
   End
   Begin VB.Menu mnuFIle 
      Caption         =   "&File"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvert 
         Caption         =   "Convert"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuSearch 
         Caption         =   "Search"
      End
      Begin VB.Menu mnuIndex 
         Caption         =   "Add/Index CD"
      End
      Begin VB.Menu mnuConfig 
         Caption         =   "Configure Settings"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuHorizon 
         Caption         =   "Tile Horizontally"
      End
      Begin VB.Menu mnuVertical 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu mnuCascade 
         Caption         =   "Cascade"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CurrentIndex%


Private Sub cmdAlign_Click(Index As Integer)

    Select Case Index
    
    Case 0
    Picture1.Align = 4
    Config.iAlign = 4
    SaveConfig
    
    Case 1
    Picture1.Align = 3
    Config.iAlign = 3
    SaveConfig
    
    End Select
    
    Picture1.SetFocus
    
End Sub

Private Sub imgTool_Click(Index As Integer)

    Select Case Index
    
    Case 0
        DoEvents
        For X = 1 To Config.iSearchWindows
            MakeSearchWindow
        Next
        
    Case 1
        With frmCatalog
            If Config.iInterface = 1 Then .WindowState = 0
            .Show
            .ZOrder
        End With
        
    Case 2
        With frmConfig
            If Config.iInterface = 1 Then .WindowState = 0
            .Show
            .ZOrder
        End With
        
    End Select

    
    
End Sub

Private Sub imgTool_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    Static IndexIsSet As Boolean
    
    If IndexIsSet = True Then
        If CurrentIndex = Index Then Exit Sub
    End If
    
    lblHighlight.Top = imgTool(Index).Top - 10
    lblHighlight.Visible = True
    CurrentIndex = Index
    IndexIsSet = True
    
End Sub


Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.Picture1.Visible = False
    
    With frmToolbar
        .Move Me.Left + 120, Me.Top + 1000
        .Show
    End With
    
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Label1_MouseDown Button, Shift, X, Y
    
End Sub

Private Sub lblHighlight_Click()

    imgTool_Click CurrentIndex
    
End Sub



Private Sub lblTool_Click(Index As Integer)

    imgTool_Click Index
    
    
End Sub

Private Sub lblTool_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgTool_MouseMove Index, Button, Shift, X, Y
    
End Sub

Private Sub MDIForm_Load()

    '// SET GLOBAL VARIABLES                        //
    cfgFile = CurDir & "\cfg.cfg"
    ExeDir = CurDir & "\"
    
    '// Check cfg file to see if it is              //
    '// initialized. If it is, then use             //
    '// the data contained there. Otherwise,        //
    '// Assign default values                       //
    
    modPublic.GetConfig
    If Config.Initialized <> True Then
        '// SET UP DEFAULTS
        With Config
            .DBFileName = "mycatalog.idx"
            .Initialized = True
            .MaxList = 20
            .wHeight = frmMDI.Height
            .wLeft = frmMDI.Left
            .wState = frmMDI.WindowState
            .wTop = frmMDI.Top
            .wWidth = frmMDI.Width
            .ListviewStyle = "3"
            .sAlbum = 1
            .sArtist = 1
            .sFilename = 1
            .sGenre = 1
            .sTitle = 1
            .sYear = 1
            .iSearchWindows = 1
            For X = 1 To 9
            .chWidth(X) = 1200
            Next
        End With
        
        modPublic.SaveConfig
    Else
        
        If Config.wState = 2 Then
            frmMDI.WindowState = 2
        Else
            frmMDI.Move Config.wLeft, Config.wTop, _
                        Config.wWidth, Config.wHeight
        End If
    
        If Config.iAlign = 4 Then Picture1.Align = 4
        
    End If
    
    '// IS DEFAULT DRIVE SET?
    If InStr(1, Config.iDefaultPlayerDrive, "\") = 0 Then
        Config.iDefaultPlayerDrive = "d:\"
        SaveConfig
    End If
    
    CatalogFileName = Trim$(Config.DBFileName)
    lblTool(0).Caption = "Search (" & Config.iSearchWindows & ")"
    
    If Config.iSearchAtStartup = 1 Then MakeSearchWindow
    
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    '// STORE FORM ATTRIBUTES TO CFG FILE               //
    With frmMDI
        Config.wHeight = .Height
        Config.wTop = .Top
        Config.wLeft = .Left
        Config.wWidth = .Width
        Config.wState = .WindowState
    End With
    
    SaveConfig
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)

    End
    
End Sub

Private Sub mnuAbout_Click()

    MsgBox "Disk Pro", vbInformation, "About Disk Pro"
         
End Sub

Private Sub mnuCascade_Click()

    Arrange vbCascade
    
End Sub

Private Sub mnuConfig_Click()

    imgTool_Click 2
    
End Sub


Private Sub mnuExit_Click()

    Unload Me
    End
    
End Sub

Private Sub mnuHorizon_Click()

    Arrange vbHorizontal
    
End Sub

Private Sub mnuIndex_Click()

    imgTool_Click 1
    
End Sub

Private Sub mnuSearch_Click()

    imgTool_Click 0
    
End Sub

Private Sub mnuVertical_Click()

    Arrange vbVertical
    
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblHighlight.Visible = False
    CurrentIndex = 99
    
End Sub

Public Sub MakeSearchWindow()

    Dim frmDBSearch As New frmSearch
    With frmDBSearch
        .Caption = "Search"
        .Show
        If Config.iInterface = 1 Then .WindowState = 0
    End With

End Sub
