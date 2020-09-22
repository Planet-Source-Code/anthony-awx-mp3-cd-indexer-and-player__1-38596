VERSION 5.00
Begin VB.Form frmToolbar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4485
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   1770
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00B49800&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4425
      Left            =   30
      ScaleHeight     =   295
      ScaleMode       =   0  'User
      ScaleWidth      =   45.327
      TabIndex        =   0
      Top             =   30
      Width           =   1710
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   1200
         Top             =   2685
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
         TabIndex        =   2
         Top             =   75
         Visible         =   0   'False
         Width           =   225
      End
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
         TabIndex        =   1
         Top             =   75
         Visible         =   0   'False
         Width           =   225
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   15
         TabIndex        =   11
         Top             =   4410
         Width           =   1695
      End
      Begin VB.Image imgTool 
         Height          =   480
         Index           =   0
         Left            =   607
         MouseIcon       =   "frmToolbar.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmToolbar.frx":030A
         Top             =   570
         Width           =   480
      End
      Begin VB.Image imgTool 
         Height          =   480
         Index           =   1
         Left            =   607
         MouseIcon       =   "frmToolbar.frx":074C
         MousePointer    =   99  'Custom
         Picture         =   "frmToolbar.frx":0A56
         Top             =   1725
         Width           =   480
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
         MouseIcon       =   "frmToolbar.frx":0E98
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   1035
         Width           =   615
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
         MouseIcon       =   "frmToolbar.frx":11A2
         MousePointer    =   99  'Custom
         TabIndex        =   8
         Top             =   2280
         Width           =   765
      End
      Begin VB.Image imgTool 
         Height          =   480
         Index           =   2
         Left            =   607
         MouseIcon       =   "frmToolbar.frx":14AC
         MousePointer    =   99  'Custom
         Picture         =   "frmToolbar.frx":17B6
         Top             =   2940
         Width           =   480
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
         MouseIcon       =   "frmToolbar.frx":1BF8
         MousePointer    =   99  'Custom
         TabIndex        =   7
         Top             =   3495
         Width           =   825
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
         TabIndex        =   5
         Top             =   90
         Width           =   645
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Height          =   2.45745e5
         Index           =   0
         Left            =   1695
         TabIndex        =   4
         Top             =   -495
         Width           =   375
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Height          =   2.45745e5
         Index           =   1
         Left            =   -360
         TabIndex        =   3
         Top             =   0
         Width           =   375
      End
      Begin VB.Label lblHighlight 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   195
         MouseIcon       =   "frmToolbar.frx":1F02
         MousePointer    =   99  'Custom
         TabIndex        =   10
         Top             =   405
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         Height          =   345
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   1635
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00715C00&
      Height          =   285
      Index           =   3
      Left            =   -15
      TabIndex        =   15
      Top             =   4455
      Width           =   2445
   End
   Begin VB.Label Label5 
      BackColor       =   &H00715C00&
      Height          =   9210
      Index           =   2
      Left            =   1740
      TabIndex        =   14
      Top             =   30
      Width           =   360
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Index           =   1
      Left            =   -555
      TabIndex        =   13
      Top             =   -405
      Width           =   2325
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Height          =   9210
      Index           =   0
      Left            =   -330
      TabIndex        =   12
      Top             =   -2775
      Width           =   360
   End
End
Attribute VB_Name = "frmToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CurrentIndex%

Private Sub Form_Load()

    SetWindowPos Me.hWnd, -1, frmMDI.Left + 240, frmMDI.Top + 8000, 1800 / 15, 4515 / 15, &H2
    doApplyTranslucency Me
    
End Sub
    

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
    Timer1_Timer
    
End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
    Timer1_Timer
    
End Sub

Private Sub Timer1_Timer()

    If Me.Left + Me.Width > (Screen.Width - 300) Then
        Me.Left = Screen.Width - Me.Width
    End If
        
    If Top < 120 Then Top = 0
    
    If Me.Left < frmMDI.Left Then
        frmMDI.Picture1.Visible = True
        Unload Me
    End If

End Sub


Private Sub imgTool_Click(Index As Integer)

    Select Case Index
    
    Case 0
        DoEvents
        For X = 1 To Config.iSearchWindows
            frmMDI.MakeSearchWindow
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

    If frmMDI.WindowState = 1 Then frmMDI.WindowState = 0
    
    
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

Private Sub lblHighlight_Click()

    imgTool_Click CurrentIndex
    
End Sub

Private Sub lblTool_Click(Index As Integer)

    imgTool_Click Index
    
End Sub

Private Sub lblTool_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    imgTool_MouseMove Index, Button, Shift, X, Y
    
End Sub

Private Sub mnuConfig_Click()

    imgTool_Click 2
    
End Sub

Private Sub mnuIndex_Click()

    imgTool_Click 1
    
End Sub

Private Sub mnuSearch_Click()

    imgTool_Click 0
    
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblHighlight.Visible = False
    CurrentIndex = 99
    
End Sub

