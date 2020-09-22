VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{22D6F304-B0F6-11D0-94AB-0080C74C7E95}#1.0#0"; "msdxm.ocx"
Begin VB.Form frmPlayer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Play MP3"
   ClientHeight    =   3315
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4935
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   4935
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Interval        =   250
      Left            =   1665
      Top             =   3900
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   930
      Left            =   270
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   21
      Text            =   "Form1.frx":0442
      Top             =   2025
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Enabled         =   0   'False
      Height          =   1125
      Left            =   180
      TabIndex        =   20
      Top             =   1935
      Width           =   4530
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "g"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   " Stop "
      Top             =   1455
      Width           =   390
   End
   Begin VB.CommandButton CmdPause 
      Caption         =   "||"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3945
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   " Pause "
      Top             =   1455
      Width           =   390
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   210
      ScaleHeight     =   270
      ScaleWidth      =   4425
      TabIndex        =   18
      Top             =   1155
      Width           =   4425
      Begin VB.Label lblSongTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Song"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   90
         TabIndex        =   19
         Top             =   30
         UseMnemonic     =   0   'False
         Width           =   420
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   240
      ScaleHeight     =   390
      ScaleWidth      =   3195
      TabIndex        =   11
      Top             =   1515
      Width           =   3195
      Begin MSComctlLib.Slider Slider1 
         Height          =   300
         Left            =   45
         TabIndex        =   12
         Top             =   75
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
         _Version        =   393216
         TickStyle       =   3
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   300
         Left            =   1530
         TabIndex        =   13
         Top             =   75
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   529
         _Version        =   393216
         Max             =   2500
         SelStart        =   2500
         TickStyle       =   3
         Value           =   2500
      End
      Begin VB.Label Label 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Vol"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2415
         TabIndex        =   16
         Top             =   75
         Width           =   465
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   2715
         TabIndex        =   15
         Top             =   75
         Width           =   435
      End
      Begin VB.Label txtTime 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   885
         TabIndex        =   14
         Top             =   75
         Width           =   330
      End
   End
   Begin VB.CommandButton CmdPlay 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   12
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   " Play "
      Top             =   1455
      Width           =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   870
      Top             =   3450
   End
   Begin VB.CommandButton Command2 
      Enabled         =   0   'False
      Height          =   495
      Left            =   180
      TabIndex        =   10
      Top             =   1455
      Width           =   3315
   End
   Begin VB.CommandButton Command3 
      Enabled         =   0   'False
      Height          =   345
      Left            =   180
      TabIndex        =   17
      Top             =   1125
      Width           =   4530
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   4275
      Picture         =   "Form1.frx":0451
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Play MP3"
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
      Index           =   1
      Left            =   195
      TabIndex        =   9
      Top             =   195
      Width           =   1155
   End
   Begin VB.Label Text1 
      BackStyle       =   0  'Transparent
      Caption         =   "Song"
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
      Height          =   630
      Left            =   480
      TabIndex        =   8
      Top             =   3960
      Width           =   4185
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   45
      Left            =   -30
      TabIndex        =   6
      Top             =   885
      Width           =   1.00005e5
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Play MP3"
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
      Index           =   0
      Left            =   225
      TabIndex        =   5
      Top             =   225
      Width           =   1155
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   15
      Left            =   195
      TabIndex        =   4
      Top             =   555
      Width           =   1.00005e5
   End
   Begin MediaPlayerCtl.MediaPlayer MediaPlayer1 
      Height          =   375
      Left            =   405
      TabIndex        =   3
      Top             =   3465
      Visible         =   0   'False
      Width           =   375
      AudioStream     =   -1
      AutoSize        =   0   'False
      AutoStart       =   -1  'True
      AnimationAtStart=   -1  'True
      AllowScan       =   -1  'True
      AllowChangeDisplaySize=   -1  'True
      AutoRewind      =   0   'False
      Balance         =   0
      BaseURL         =   ""
      BufferingTime   =   5
      CaptioningID    =   ""
      ClickToPlay     =   -1  'True
      CursorType      =   0
      CurrentPosition =   -1
      CurrentMarker   =   0
      DefaultFrame    =   ""
      DisplayBackColor=   0
      DisplayForeColor=   16777215
      DisplayMode     =   0
      DisplaySize     =   4
      Enabled         =   -1  'True
      EnableContextMenu=   -1  'True
      EnablePositionControls=   -1  'True
      EnableFullScreenControls=   0   'False
      EnableTracker   =   -1  'True
      Filename        =   ""
      InvokeURLs      =   -1  'True
      Language        =   -1
      Mute            =   0   'False
      PlayCount       =   1
      PreviewMode     =   0   'False
      Rate            =   1
      SAMILang        =   ""
      SAMIStyle       =   ""
      SAMIFileName    =   ""
      SelectionStart  =   -1
      SelectionEnd    =   -1
      SendOpenStateChangeEvents=   -1  'True
      SendWarningEvents=   -1  'True
      SendErrorEvents =   -1  'True
      SendKeyboardEvents=   0   'False
      SendMouseClickEvents=   0   'False
      SendMouseMoveEvents=   0   'False
      SendPlayStateChangeEvents=   -1  'True
      ShowCaptioning  =   0   'False
      ShowControls    =   -1  'True
      ShowAudioControls=   -1  'True
      ShowDisplay     =   0   'False
      ShowGotoBar     =   0   'False
      ShowPositionControls=   -1  'True
      ShowStatusBar   =   0   'False
      ShowTracker     =   -1  'True
      TransparentAtStart=   0   'False
      VideoBorderWidth=   0
      VideoBorderColor=   0
      VideoBorder3D   =   0   'False
      Volume          =   0
      WindowlessVideo =   0   'False
   End
   Begin VB.Image imgLogo 
      Height          =   885
      Left            =   1350
      Picture         =   "Form1.frx":0893
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3585
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      Height          =   915
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   1.00005e5
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tinseconden As Integer
Dim minuten As Integer
Dim seconden As Integer

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Unicode nul terminated string
Const REG_DWORD = 4                      ' 32-bit number

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long



Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    
End Function



Private Sub CmdPause_Click()
    
    Static Paused As Boolean
    
    If Paused = True Then
        Paused = False
        MediaPlayer1.Play
        
    Else
        Paused = True
        MediaPlayer1.Pause

    End If

    Text2.SetFocus
    
End Sub

Public Sub CmdPlay_Click()

    On Error Resume Next
    
    MediaPlayer1.Filename = Replace(Text1.Caption, "\\", "\")
    
    If Text1.Caption <> "" Then
        MediaPlayer1.Play
        Slider1.Max = MediaPlayer1.Duration
        CmdPause.Enabled = True

    Else
        MsgBox "No file to play", vbOKOnly, "Error"
    
    End If

    Text2.SetFocus
    
End Sub

Private Sub CmdStop_Click()

    MediaPlayer1.Stop
    Slider1.Value = 0
    CmdPause.Enabled = False

    Text2.SetFocus
    
End Sub

Private Sub Form_Load()

    SetWindowPos Me.hWnd, -1, frmMDI.Left + 3000, frmMDI.Top + 8000, 5025 / 15, 3690 / 15, &H2
    doApplyTranslucency Me

End Sub


Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
 
End Sub

Private Sub imgLogo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
    
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
 
End Sub

Private Sub Label7_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
 
End Sub

Private Sub MediaPlayer1_EndOfStream(ByVal Result As Long)

    MediaPlayer1.Stop

End Sub

Private Sub Slider1_Scroll()

    MediaPlayer1.CurrentPosition = Slider1.Value
    Slider1.ToolTipText = ""
    
End Sub

Private Sub Slider3_Scroll()

    sha% = Slider3.Value - 2500
    MediaPlayer1.Volume = sha
    
    On Error Resume Next
    Label2.Caption = Slider3.Value \ 25 & " %"

    Slider3.ToolTipText = ""
    
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Private Sub Timer1_Timer()

    Slider1.Value = MediaPlayer1.CurrentPosition
    tinseconden = MediaPlayer1.CurrentPosition
    Dim min As Integer
    Dim sec As Integer
    min = tinseconden \ 60
    sec = tinseconden - (min * 60)
    If sec = "-1" Then sec = "0"
    txtTime.Caption = min & ":" & Right$("00" & Trim$(Str$(sec)), 2)

End Sub


Private Sub Timer2_Timer()

    Static Initialized As Boolean
    Static FullSongName$
    Static ScrollForward As Boolean
    Static ScrollCount%
    
    If Initialized <> True Then
        FullSongName = Left$(lblSongTitle.Caption & String(42, " "), 42)
        Initialized = True
        Exit Sub
    End If
    
    ScrollCount = ScrollCount + 1
    
    If ScrollCount < 43 Then
    
        FullSongName = Right$(FullSongName, Len(FullSongName) - 1) _
                     & Left$(FullSongName, 1)
                 
        lblSongTitle.Caption = FullSongName
        
    Else
        
        FullSongName = Right$(FullSongName, 1) _
                     & Left$(FullSongName, Len(FullSongName) - 1)

        lblSongTitle.Caption = FullSongName
        
    End If
        
    If ScrollCount >= 84 Then ScrollCount = 0
        
End Sub
