VERSION 5.00
Begin VB.Form frmCatalog 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Catalog Pro"
   ClientHeight    =   8760
   ClientLeft      =   165
   ClientTop       =   165
   ClientWidth     =   10260
   Icon            =   "disksearch.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8760
   ScaleWidth      =   10260
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H00B49800&
      ForeColor       =   &H80000008&
      Height          =   1590
      Left            =   1755
      ScaleHeight     =   1560
      ScaleWidth      =   3615
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   3645
      Begin VB.Timer Timer1 
         Interval        =   700
         Left            =   600
         Top             =   1005
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2280
         TabIndex        =   18
         Top             =   960
         Width           =   1155
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Height          =   1800
         Index           =   3
         Left            =   3555
         TabIndex        =   25
         Top             =   -90
         Width           =   555
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Height          =   1800
         Index           =   2
         Left            =   -495
         TabIndex        =   24
         Top             =   -180
         Width           =   555
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Height          =   105
         Index           =   1
         Left            =   -135
         TabIndex        =   23
         Top             =   1500
         Width           =   3930
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Height          =   105
         Index           =   0
         Left            =   -180
         TabIndex        =   22
         Top             =   -45
         Width           =   3930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00715C00&
         BorderWidth     =   3
         X1              =   780
         X2              =   3405
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   195
         Picture         =   "disksearch.frx":0442
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lblActive 
         Caption         =   "false"
         Height          =   300
         Left            =   165
         TabIndex        =   21
         Top             =   1905
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblCancelFlag 
         Caption         =   "0"
         Height          =   285
         Left            =   2790
         TabIndex        =   20
         Top             =   1920
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Insert CD"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   765
         TabIndex        =   19
         Top             =   330
         Width           =   2115
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4725
      ScaleHeight     =   285
      ScaleWidth      =   1530
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   975
      Width           =   1560
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   15
         ScaleHeight     =   240
         ScaleWidth      =   1470
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   15
         Width           =   1470
         Begin VB.DriveListBox drvList 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   315
            Left            =   -30
            TabIndex        =   15
            Top             =   -30
            Width           =   1545
         End
      End
   End
   Begin VB.CheckBox chkAuto 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Auto-Increment Volume Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   225
      Left            =   720
      TabIndex        =   5
      Top             =   1320
      Value           =   1  'Checked
      Width           =   2505
   End
   Begin VB.ListBox lstFoundFiles 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   6435
      IntegralHeight  =   0   'False
      Left            =   3465
      TabIndex        =   4
      Top             =   2115
      Width           =   6630
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Catalog"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6450
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   975
      Width           =   1140
   End
   Begin VB.TextBox txtVolName 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      Left            =   720
      TabIndex        =   2
      Top             =   975
      Width           =   2520
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   315
      Left            =   3390
      TabIndex        =   1
      Text            =   "*.mp3"
      Top             =   975
      Width           =   1215
   End
   Begin VB.DirListBox dirList 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   6165
      Left            =   120
      TabIndex        =   0
      Top             =   2130
      Width           =   3255
   End
   Begin VB.FileListBox filList 
      Height          =   2040
      Left            =   705
      TabIndex        =   16
      Top             =   4830
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List of files Indexed"
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
      Left            =   3465
      TabIndex        =   28
      Top             =   1875
      Width           =   1650
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Folders in Selected Drive"
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
      Left            =   135
      TabIndex        =   27
      Top             =   1875
      Width           =   2085
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Drive"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   4755
      TabIndex        =   26
      Top             =   720
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "disksearch.frx":0884
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   15
      Left            =   705
      TabIndex        =   10
      Top             =   540
      Width           =   1.00005e5
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Index a CD"
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
      TabIndex        =   9
      Top             =   165
      Width           =   1365
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   45
      Left            =   0
      TabIndex        =   8
      Top             =   1650
      Width           =   1.00005e5
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Volume Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "File Types"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3405
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Index a CD"
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
      TabIndex        =   11
      Top             =   180
      Width           =   1365
   End
   Begin VB.Image imgLogo 
      Height          =   1650
      Left            =   0
      Picture         =   "disksearch.frx":0CC6
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3570
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Height          =   1695
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   1.00005e5
   End
End
Attribute VB_Name = "frmCatalog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' //////////////////////////////////////////////////////////////////////////////////////
'
Private Type IconeTray
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Dim IconeT As IconeTray

Private Type DrvInfo
    DriveLetter As String
    DriveName As String
    ID As Long
    CDOpen As Boolean
End Type

Private Const AJOUT = &H0
Private Const MODIF = &H1
Private Const SUPPRIME = &H2
Private Const MOUSEMOVE = &H200
Private Const MESSAGE = &H1
Private Const Icone = &H2
Private Const TIP = &H4

Private Const DOUBLE_CLICK_GAUCHE = &H203
Private Const BOUTON_GAUCHE_POUSSE = &H201
Private Const BOUTON_GAUCHE_LEVE = &H202
Private Const DOUBLE_CLICK_DROIT = &H206
Private Const BOUTON_DROIT_POUSSE = &H204
Private Const BOUTON_DROIT_LEVE = &H205

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As IconeTray) As Boolean
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long

Dim APICheck1 As Boolean
Dim CDDrives() As DrvInfo
'
' /////////////////////////////////////////////////////////////////////////////////////

Private Sub dirbox1_Change()

    File1.Path = dir1.Path

End Sub

Private Sub drvbox1_Change()

    On Error GoTo drivehandler
    
    'if new drive was selected, the dir list box updates display]
        dir1.Path = Drive1.Drive
        Exit Sub
    
    'if there is an error reset drive1.drive with
    'the drive from dir1.path
    
drivehandler:
    Drive1.Drive = dir1.Path
    Exit Sub
    
End Sub

Private Sub dir1_Change()

    File1.Path = dir1.Path

End Sub

Private Sub chkAuto_Click()

    Select Case chkAuto.Value
    
    Case Checked
        txtVolName.Enabled = False
        txtVolName.Text = ""

    Case Unchecked
        txtVolName.Enabled = True

    End Select
    
End Sub

Private Sub Command1_Click()

    MousePointer = 11
    
'// AUTO INCREMENT?
    If chkAuto.Value = Checked Then
    
        highVal% = 0
        
        '// CHECK FOR HIGHEST VAL
        ff = FreeFile
        Open CatalogFileName For Random As ff Len = Len(IndexFile)
        nRecs = LOF(ff) / Len(IndexFile)
        For X = 1 To nRecs
            Get ff, X, IndexFile
            tmpVal% = Val(Trim$(IndexFile.VolumeName))
            If tmpVal% > highVal% Then highVal% = tmpVal%
        Next
        Close ff
    
        '// PLACE VOL NAME INTO TEXT FIELD
        txtVolName.Text = Trim$(Str$(tmpVal + 1))
    
    ElseIf txtVolName = "" Then
        
        MsgBox "Can not proceed - Volume Name Field is blank.", vbCritical
        MousePointer = 0
        Exit Sub
    
    End If
    
    On Error GoTo DrError
    
    Dim DLetter$
    DLetter$ = Left$(drvList.Drive, 1)
    
    DriveAvailable = CDEjector.GetDriveState(DLetter)
    
    If DriveAvailable = False Then
        CDEjector.openCD (DLetter)
        MsgBox "Insert CD into Drive " & DLetter & ", then click OK.", vbInformation, "Insert CD"
        CDEjector.closeCD (DLetter)
        drvList_Change
    End If
    
    MousePointer = 11
    
    drvList_Change
    
'// DOES THIS VOLUME NAME EXIST?
    ff = FreeFile
    Open CatalogFileName For Random As ff Len = Len(IndexFile)
    nRecs = LOF(ff) / Len(IndexFile)
    For X = 1 To nRecs
    Get ff, 1, IndexFile
    If Trim(LCase(txtVolName.Text)) = Trim(LCase(IndexFile.VolumeName)) Then
        MsgBox "Volume Name Already in use."
        Exit Sub
    End If
    Next
    Close ff
    
    lstFoundFiles.Clear
    lstFoundFiles.Refresh
        
    Dim firstPath$, DirCount%
    
    If Text2 = "" Then Text2 = "*.*"
    
    filList.Pattern = Text2
    firstPath$ = dirList.Path
    DirCount = dirList.ListCount

    NumFiles = 0                       ' Reset found files indicator.
    Result = DirDiver(firstPath, DirCount, "")
    filList.Path = dirList.Path
    
    
' *******************************************************************
' SAVE TO FILE

    MousePointer = 11
    
    ff = FreeFile
    Open CatalogFileName For Random As ff Len = Len(IndexFile)
    nRecs = LOF(ff) / Len(IndexFile)
    startrecord% = nRecs + 1
    Dim thisIDTag As New idtag
    
    For X = 0 To (lstFoundFiles.ListCount - 1)
        
        IndexFile.VolumeName = txtVolName
        tmp$ = Trim$(lstFoundFiles.List(X))
        IndexFile.Filename = tmp$
        
        '// ID3 INFO
        thisIDTag.Filename = Me.dirList.Path & "\" & tmp$
        With IndexFile.ID3
            .Album = thisIDTag.Album
            .Artist = thisIDTag.Artist
            .Comment = thisIDTag.Comment
            .Genre = thisIDTag.Genre
            .Title = thisIDTag.Title
            .Year = thisIDTag.Year
        End With
        
        '// MP3 Info
        
        Dim accMP3Info As Mp3Info
        getMP3Info thisIDTag.Filename, accMP3Info
        With IndexFile.Mp3Info
            .Size = accMP3Info.Size
            .Length = accMP3Info.Length
            .Layer = accMP3Info.MPEG & " " & accMP3Info.Layer
            .BitRate = accMP3Info.BitRate
            .FreqChannel = accMP3Info.FREQ & " " & accMP3Info.CHANNELS
            .CRC = accMP3Info.CRC
            .Copy = accMP3Info.COPYRIGHT
            .Emphasis = accMP3Info.Emphasis
            .Original = accMP3Info.Original
        End With
        
        Put ff, startrecord, IndexFile
        startrecord = startrecord + 1
        
    Next
    
    Close ff
    
    CDEjector.openCD DLetter
    
    MousePointer = 0
    
    answer = MsgBox("Would you like to index another CD in the same drive?", _
                    vbQuestion + vbYesNo, "Index Another CD?")
                    
    If answer = vbYes Then GoTo AnotherOne
    
    Exit Sub
    
DrError:
    MsgBox "Drive Not Ready." & vbCrLf & vbCrLf _
          & "Error: " & Err.Number & " - " & Err.Description, vbInformation
    MousePointer = 0
    Exit Sub
    

AnotherOne:
    
    Me.lstFoundFiles.Clear
    Me.Picture3.Visible = True
    lblActive.Caption = "false"
    
    ' DoEvents
    Do Until lblActive.Caption = "true"
        
        DoEvents
        If lblCancelFlag.Caption = "1" Then
            lblCancelFlag.Caption = "0"
            Picture3.Visible = False
            Command1.Enabled = False
            Exit Sub
        End If
        CheckActive
    
    Loop
    
    Label3.Visible = True
    Label3.Refresh
    Picture3.Refresh
    
    Picture3.Visible = False
    drvList_Change
    Command1_Click
    

End Sub

Private Sub Drive1_Change()

    On Error GoTo drivehandler
    'if new drive was selected, the dir list box updates display]
    dir1.Path = Drive1.Drive
    Exit Sub
    
    'if there is an error reset drive1.drive with
    'the drive from dir1.path
    
drivehandler:
    Drive1.Drive = dir1.Path
    Exit Sub
    
End Sub


Private Sub Command2_Click()

    lblCancelFlag.Caption = "1"

End Sub

Private Sub dirList_Change()
    
    Me.filList.Path = dirList.Path
    
End Sub

Private Sub drvList_Change()
    
    On Error GoTo Err
    
    lblActive.Caption = "false"
    Command1.Enabled = True
    MousePointer = 13
    dirList.Path = drvList.Drive
    dirList.Refresh
    MousePointer = 0
    lblActive.Caption = "true"
    
    Exit Sub
    
Err:
    MousePointer = 0

End Sub


Private Function DirDiver(NewPath As String, DirCount As Integer, BackUp As String) As Integer

    Static FirstErr As Integer
    Dim DirsToPeek As Integer, AbandonSearch As Integer, ind As Integer
    Dim OldPath As String, ThePath As String, entry As String
    Dim retval As Integer
    
    SearchFlag = True           ' Set flag so user can interrupt.
    DirDiver = False            ' Set to True if error.
    retval = DoEvents()         ' Check for events
    
    If SearchFlag = False Then
        DirDiver = True
        Exit Function
    End If
    
    
    DirsToPeek = dirList.ListCount                  ' How many directories below this?
    Do While DirsToPeek > 0 And SearchFlag = True
        OldPath = dirList.Path                      ' Save old path for next recursion.
        dirList.Path = NewPath
        If dirList.ListCount > 0 Then
            ' Get to the node bottom.
            dirList.Path = dirList.List(DirsToPeek - 1)
            AbandonSearch = DirDiver((dirList.Path), DirCount%, OldPath)
        End If
        ' Go up one level in directories.
        DirsToPeek = DirsToPeek - 1
        If AbandonSearch = True Then Exit Function
    Loop
    ' Call function to enumerate files.
    
    If filList.ListCount Then
        If Len(dirList.Path) <= 3 Then             ' Check for 2 bytes/character
            ThePath = dirList.Path                  ' If at root level, leave as is...
        Else
            ThePath = dirList.Path + "\"            ' Otherwise put "\" before the filename.
        End If
        For ind = 0 To filList.ListCount - 1        ' Add conforming files in this directory to the list box.
            entry = ThePath + filList.List(ind)
            lstFoundFiles.AddItem Right$(entry, Len(entry) - 2)
            'MsgBox Str(Val(lblCount.Caption) + 1)
        Next ind
    End If
    
    If BackUp <> "" Then        ' If there is a superior directory, move it.
        dirList.Path = BackUp
    End If
    
    Exit Function


DirDriverHandler:
    If Err = 7 Then             ' If Out of Memory error occurs, assume the list box just got full.
        DirDiver = True         ' Create Msg and set return value AbandonSearch.
        MsgBox "You've filled the list box. Abandoning search..."
        Exit Function           ' Note that the exit procedure resets Err to 0.
    Else                        ' Otherwise display error message and quit.
        MsgBox Error
        End
    End If

End Function

Private Sub Form_Load()

    On Error Resume Next
    Me.drvList.Drive = Config.iDefaultPlayerDrive
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    dirList.Height = ScaleHeight - dirList.Top - 180
    lstFoundFiles.Height = ScaleHeight - lstFoundFiles.Top - 180
    lstFoundFiles.Width = ScaleWidth - lstFoundFiles.Left - 180
    
    With imgLogo
        .Top = 15
        .Left = ScaleWidth - .Width - 15
    End With
    
End Sub



Private Sub Text2_Change()

    doUpdateFileBox
    
End Sub

Private Sub Text2_Click()

    doUpdateFileBox
    
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)

    doUpdateFileBox
    
End Sub

Private Sub doUpdateFileBox()

    On Error GoTo doDefault
    filList.Pattern = Text2
    filList.Refresh
        
    Exit Sub
       
doDefault:
    filList.Pattern = "*.*"
    
End Sub

Private Sub Text4_Change()
    
    Command2.Default = True
    
End Sub


Private Sub Timer1_Timer()

    Label3.Visible = Not Label3.Visible
    
End Sub

Private Sub txtVolName_Change()

    Command1.Default = True
    
End Sub

Private Sub txtVolName_GotFocus()

    Command1.Default = True
    
End Sub

Private Sub CheckActive()

    On Error GoTo Err
    
    lblActive.Caption = "false"
    'Command1.Enabled = True
    MousePointer = 13
    dirList.Path = drvList.Drive
    dirList.Refresh
    MousePointer = 0
    lblActive.Caption = "true"
    
    Exit Sub
    
Err:
    MousePointer = 0
    
End Sub


