VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSearch 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form2"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10725
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   9120
   ScaleWidth      =   10725
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picDeleteProgress 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFDE91&
      ForeColor       =   &H80000008&
      Height          =   1665
      Left            =   1530
      ScaleHeight     =   1635
      ScaleWidth      =   4005
      TabIndex        =   51
      Top             =   2985
      Visible         =   0   'False
      Width           =   4035
      Begin VB.CommandButton Command4 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2805
         TabIndex        =   60
         Top             =   1140
         Width           =   855
      End
      Begin MSComctlLib.ProgressBar pbDelete 
         Height          =   270
         Left            =   360
         TabIndex        =   52
         Top             =   810
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   476
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblActivityTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Removing Selected Items from List"
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
         Left            =   150
         TabIndex        =   53
         Top             =   135
         Width           =   2985
      End
      Begin VB.Label Label24 
         BackColor       =   &H00B49800&
         Height          =   360
         Left            =   45
         TabIndex        =   59
         Top             =   45
         Width           =   3900
      End
      Begin VB.Label lblCounter 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   360
         TabIndex        =   58
         Top             =   510
         Width           =   135
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         Height          =   1800
         Index           =   3
         Left            =   0
         TabIndex        =   57
         Top             =   -180
         Width           =   30
      End
      Begin VB.Label Label22 
         BackColor       =   &H00B49800&
         Height          =   1800
         Index           =   2
         Left            =   3975
         TabIndex        =   56
         Top             =   15
         Width           =   30
      End
      Begin VB.Label Label22 
         BackColor       =   &H00B49800&
         Height          =   30
         Index           =   1
         Left            =   -15
         TabIndex        =   55
         Top             =   1605
         Width           =   4455
      End
      Begin VB.Label Label22 
         BackColor       =   &H00F8F8F8&
         Height          =   30
         Index           =   0
         Left            =   -135
         TabIndex        =   54
         Top             =   0
         Width           =   4455
      End
   End
   Begin VB.PictureBox picSearchlist 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFDE91&
      ForeColor       =   &H80000008&
      Height          =   3660
      Left            =   705
      ScaleHeight     =   3630
      ScaleWidth      =   2535
      TabIndex        =   37
      Top             =   1245
      Visible         =   0   'False
      Width           =   2565
      Begin VB.CommandButton Command3 
         Caption         =   "Clear"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   1620
         TabIndex        =   44
         Top             =   75
         Width           =   660
      End
      Begin VB.CommandButton Command2 
         Caption         =   "x"
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
         Left            =   2295
         TabIndex        =   43
         Top             =   75
         Width           =   225
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3255
         IntegralHeight  =   0   'False
         Left            =   -15
         Sorted          =   -1  'True
         TabIndex        =   39
         Top             =   390
         Width           =   2565
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keyword History"
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
         TabIndex        =   40
         Top             =   90
         Width           =   1395
      End
      Begin VB.Label Label18 
         BackColor       =   &H00B49800&
         Height          =   360
         Left            =   15
         TabIndex        =   41
         Top             =   15
         Width           =   2505
      End
   End
   Begin VB.PictureBox picKeywordHistoryShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   3660
      Left            =   735
      ScaleHeight     =   3630
      ScaleWidth      =   2535
      TabIndex        =   42
      Top             =   1275
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.CheckBox chkClearList 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Clear List on Search"
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
      Height          =   210
      Left            =   2475
      TabIndex        =   50
      Top             =   1320
      Width           =   1830
   End
   Begin VB.PictureBox picRightClickMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFDE91&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   1710
      ScaleHeight     =   1380
      ScaleWidth      =   1905
      TabIndex        =   46
      Top             =   5400
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Label lblDelete 
         BackStyle       =   0  'Transparent
         Caption         =   "Delete from List"
         Enabled         =   0   'False
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
         Left            =   165
         TabIndex        =   62
         Top             =   870
         Width           =   1635
      End
      Begin VB.Label lblPlay 
         BackStyle       =   0  'Transparent
         Caption         =   "Play Song"
         Enabled         =   0   'False
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
         Left            =   150
         TabIndex        =   61
         Top             =   495
         Width           =   1500
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Context Menu"
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
         TabIndex        =   47
         Top             =   75
         Width           =   1185
      End
      Begin VB.Label Label20 
         BackColor       =   &H00B49800&
         Height          =   360
         Left            =   15
         TabIndex        =   48
         Top             =   15
         Width           =   1875
      End
      Begin VB.Label lblContextHL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   45
         TabIndex        =   63
         Top             =   405
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   ".."
      Height          =   270
      Left            =   3015
      TabIndex        =   38
      Top             =   960
      Width           =   225
   End
   Begin VB.PictureBox picExportMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFDE91&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   5355
      ScaleHeight     =   1380
      ScaleWidth      =   2340
      TabIndex        =   33
      Top             =   1215
      Visible         =   0   'False
      Width           =   2370
      Begin VB.Label lblToText 
         BackStyle       =   0  'Transparent
         Caption         =   "Text"
         Enabled         =   0   'False
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
         Left            =   150
         TabIndex        =   65
         Top             =   495
         Width           =   2040
      End
      Begin VB.Label lblToHtml 
         BackStyle       =   0  'Transparent
         Caption         =   "HTML"
         Enabled         =   0   'False
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
         Left            =   165
         TabIndex        =   64
         Top             =   870
         Width           =   2040
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Export to"
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
         TabIndex        =   34
         Top             =   75
         Width           =   780
      End
      Begin VB.Label Label16 
         BackColor       =   &H00B49800&
         Height          =   360
         Left            =   15
         TabIndex        =   35
         Top             =   15
         Width           =   2310
      End
      Begin VB.Label lblExportHL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   45
         TabIndex        =   66
         Top             =   405
         Visible         =   0   'False
         Width           =   2220
      End
   End
   Begin VB.PictureBox picExportShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   5385
      ScaleHeight     =   1380
      ScaleWidth      =   2340
      TabIndex        =   36
      Top             =   1245
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.PictureBox picOptMenu 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFDE91&
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   4470
      ScaleHeight     =   3915
      ScaleWidth      =   2340
      TabIndex        =   17
      Top             =   1215
      Visible         =   0   'False
      Width           =   2370
      Begin VB.CheckBox chkVolume 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDE91&
         Caption         =   "Volume Name/Number"
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
         Height          =   270
         Left            =   360
         TabIndex        =   45
         Top             =   840
         Width           =   2040
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   360
         ScaleHeight     =   300
         ScaleWidth      =   1530
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   3330
         Width           =   1560
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   30
            ScaleHeight     =   240
            ScaleWidth      =   1470
            TabIndex        =   29
            Top             =   30
            Width           =   1470
            Begin VB.ComboBox cmbMax 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "frmSearch.frx":0442
               Left            =   -30
               List            =   "frmSearch.frx":047F
               TabIndex        =   30
               Text            =   "Combo1"
               Top             =   -30
               Width           =   1545
            End
         End
      End
      Begin VB.CheckBox chkFilename 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDE91&
         Caption         =   "Filename"
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
         Height          =   270
         Left            =   360
         TabIndex        =   23
         Top             =   555
         Width           =   1020
      End
      Begin VB.CheckBox chkAlbum 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDE91&
         Caption         =   "ID3 Tag: Album"
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
         Height          =   270
         Left            =   360
         TabIndex        =   22
         Top             =   1125
         Width           =   1590
      End
      Begin VB.CheckBox chkArtist 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDE91&
         Caption         =   "ID3 Tag: Artist"
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
         Height          =   270
         Left            =   360
         TabIndex        =   21
         Top             =   1410
         Width           =   1590
      End
      Begin VB.CheckBox chkGenre 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDE91&
         Caption         =   "ID3 Tag: Genre"
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
         Height          =   270
         Left            =   360
         TabIndex        =   20
         Top             =   1980
         Width           =   1485
      End
      Begin VB.CheckBox chkTitle 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDE91&
         Caption         =   "ID3 Tag: Title"
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
         Height          =   270
         Left            =   360
         TabIndex        =   19
         Top             =   1695
         Width           =   1380
      End
      Begin VB.CheckBox chkYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFDE91&
         Caption         =   "ID3 Tag: Year"
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
         Height          =   270
         Left            =   360
         TabIndex        =   18
         Top             =   2265
         Width           =   1485
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "# Records to Show"
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
         TabIndex        =   31
         Top             =   2880
         Width           =   1590
      End
      Begin VB.Label Label14 
         BackColor       =   &H00B49800&
         Height          =   360
         Left            =   15
         TabIndex        =   32
         Top             =   2820
         Width           =   2310
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fields to Search"
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
         TabIndex        =   24
         Top             =   90
         Width           =   1335
      End
      Begin VB.Label Label12 
         BackColor       =   &H00B49800&
         Height          =   360
         Left            =   15
         TabIndex        =   26
         Top             =   15
         Width           =   2310
      End
   End
   Begin VB.PictureBox picOptShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3945
      Left            =   4500
      ScaleHeight     =   3945
      ScaleWidth      =   2370
      TabIndex        =   16
      Top             =   1245
      Visible         =   0   'False
      Width           =   2370
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      Caption         =   "Show Large Icons"
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
      Height          =   210
      Left            =   705
      TabIndex        =   15
      Top             =   1305
      Width           =   1830
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   9270
      Top             =   5055
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   9150
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":04DD
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9885
      Top             =   8310
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSearch.frx":092F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdSearch 
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
      Height          =   345
      Left            =   3360
      TabIndex        =   1
      Top             =   930
      Width           =   885
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   705
      TabIndex        =   0
      Top             =   930
      Width           =   2565
   End
   Begin VB.PictureBox picRightClickMenuShadow 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H80000008&
      Height          =   1410
      Left            =   1740
      ScaleHeight     =   1380
      ScaleWidth      =   1905
      TabIndex        =   49
      Top             =   5430
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   7215
      Left            =   120
      TabIndex        =   3
      ToolTipText     =   " Double-Click to Play a Song "
      Top             =   1800
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   12726
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList2"
      ForeColor       =   0
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Volume"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Filename"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Album"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Artist"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Genre"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Title"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Year"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Comment"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Record"
         Object.Width           =   0
      EndProperty
      Picture         =   "frmSearch.frx":0D81
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Height          =   1530
      Left            =   390
      TabIndex        =   4
      Top             =   1785
      Width           =   99999
      Begin VB.CommandButton cmdCancel 
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
         Height          =   330
         Left            =   8550
         TabIndex        =   13
         Top             =   915
         Width           =   1035
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   285
         Left            =   165
         TabIndex        =   5
         Top             =   540
         Visible         =   0   'False
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   503
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblCancelFlag 
         Caption         =   "0"
         Height          =   285
         Left            =   8535
         TabIndex        =   14
         Top             =   1245
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   165
         TabIndex        =   7
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Performing Search ..."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   165
         TabIndex        =   6
         Top             =   285
         Width           =   4845
      End
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Export"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5370
      MouseIcon       =   "frmSearch.frx":237F
      MousePointer    =   99  'Custom
      TabIndex        =   27
      Top             =   1020
      Width           =   765
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4485
      MouseIcon       =   "frmSearch.frx":2689
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   45
      Left            =   0
      TabIndex        =   12
      Top             =   1650
      Width           =   1.00005e5
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
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
      TabIndex        =   11
      Top             =   165
      Width           =   885
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
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frmSearch.frx":2993
      Top             =   180
      Width           =   480
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search"
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
      TabIndex        =   9
      Top             =   180
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Find Word or Phrase"
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
      Height          =   195
      Left            =   735
      TabIndex        =   2
      Top             =   705
      Width           =   1470
   End
   Begin VB.Image imgLogo 
      Height          =   1650
      Left            =   7155
      Picture         =   "frmSearch.frx":2DD5
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3570
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Height          =   1695
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   1.00005e5
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

    Select Case Check1.Value
    
    Case Checked
    ListView1.View = lvwIcon
    
    Case Unchecked
    ListView1.View = lvwReport
    
    End Select
    
End Sub

Private Sub chkAlbum_Click()

    Config.sAlbum = chkAlbum.Value
    SaveConfig
    
End Sub

Private Sub chkArtist_Click()

    Config.sArtist = chkArtist.Value
    SaveConfig
    
End Sub

Private Sub chkClearList_Click()

    Config.iClearList = chkClearList.Value
    SaveConfig
    
End Sub

Private Sub chkFilename_Click()

    Config.sFilename = chkFilename.Value
    SaveConfig
    
End Sub

Private Sub chkGenre_Click()

    Config.sGenre = chkGenre.Value
    SaveConfig
    
End Sub

Private Sub chkTitle_Click()

    Config.sTitle = chkTitle.Value
    SaveConfig
    
End Sub

Private Sub chkVolume_Click()

    Config.sVolume = chkVolume.Value
    SaveConfig
    
End Sub

Private Sub chkYear_Click()

    Config.sYear = chkYear.Value
    SaveConfig
    
End Sub

Private Sub cmbMax_Change()

    Config.MaxList = Val(cmbMax.Text)
    SaveConfig
    
End Sub

Private Sub cmbMax_Click()

    Config.MaxList = Val(cmbMax.Text)
    SaveConfig
    
End Sub

Private Sub cmbMax_KeyPress(KeyAscii As Integer)

    KeyAscii = 0
    
End Sub

Private Sub cmdCancel_Click()

    lblCancelFlag.Caption = "1"

End Sub

Private Sub DeleteEntries()

    If ListView1.ListItems.Count < 1 Then Exit Sub
    
    HideRightClickMenu
    
    answer = MsgBox("Are you sure you want to remove the selected item(s)", vbQuestion + vbYesNo, "Remove?")
    
    Select Case answer
    
    Case vbYes
    gCancelProcess = False
    
    lblActivityTitle.Caption = "Removing Selected Items from List"
    
    Dim nToDelete%
    nToDelete = 0
    
    For X = 1 To ListView1.ListItems.Count
    If ListView1.ListItems(X).Selected = True Then
        ListView1.ListItems(X).Text = "___!!R"
        dcount = dcount + 1
        nToDelete = nToDelete + 1
    End If
    Next
    
    MousePointer = 13
    
    tmp = ListView1.SortKey
    ListView1.Sorted = False
    ListView1.Sorted = True
    ListView1.SortKey = 1
    ListView1.SortKey = 0
    ListView1.Refresh
    X = 1
    
    pbDelete.Visible = True
    pbDelete.min = 1
    pbDelete.Value = 1
    If dcount <> 1 Then
        pbDelete.Max = dcount
    Else
        pbDelete.Max = 2
    End If
    
    Frame1.Visible = False
    ListView1.Visible = False
    
    Me.picDeleteProgress.Visible = True
    picDeleteProgress.Refresh
    

    Do Until ListView1.ListItems(1).Text <> "___!!R"
    
        DoEvents
        ListView1.ListItems.Remove X
        rCount = rCount + 1
        Me.pbDelete.Value = rCount
        lblCounter.Caption = "Deleted " & rCount & " of " & nToDelete & " entries."
        lblCounter.Refresh
        If gCancelProcess = True Then Exit Do
        If ListView1.ListItems.Count = 0 Then Exit Do
    Loop
    
    MousePointer = 0
    ListView1.SortKey = tmp
    
    Me.picDeleteProgress.Visible = False
    
    ListView1.Visible = True
    Frame1.Visible = True
    
    End Select
    
End Sub


Private Sub HideRightClickMenu()

    Me.picRightClickMenu.Visible = False
    Me.picRightClickMenuShadow.Visible = False

End Sub
Private Sub ShowRightClickMenu()

    Me.picRightClickMenu.Visible = True
    Me.picRightClickMenuShadow.Visible = True

End Sub

Private Sub cmdSearch_Click()

    If Len(Trim$(Text4.Text)) = 0 Then
        MsgBox "Please enter a search word or phrase.", vbExclamation
        Exit Sub
    End If
    
    doAddKeyword
    
    Label4.Caption = ""
    
    MousePointer = 13
    ListView1.Visible = False
    ListView1.Sorted = False
    Refresh
    MaxVal% = cmbMax.Text
    
    With ProgressBar1
        .min = 0
        .Max = 1
        .Value = 0
        .Visible = True
    End With
    
    If chkClearList.Value = Checked Then ListView1.ListItems.Clear
    
    Label4 = ""
    matches% = 0
    
    CatalogFileName = Trim$(Config.DBFileName)
    If InStr(1, CatalogFileName, "\") > 0 Then
        CatalogFileName = ExeDir & CatalogFileName
    End If
    
    ff = FreeFile
    Open CatalogFileName For Random As ff Len = Len(IndexFile)
    nRecs = LOF(ff) / Len(IndexFile)
    ProgressBar1.Max = nRecs + 1
    For vloop% = 1 To nRecs
    DoEvents
    Get ff, vloop, IndexFile
    
    ' BUILD SEARCH STRING
    sStr$ = ""
    With IndexFile
        If chkFilename.Value = Checked Then sStr = sStr & Trim$(.Filename)
        If chkAlbum.Value = Checked Then sStr = sStr & Trim$(.ID3.Album)
        If chkArtist.Value = Checked Then sStr = sStr & Trim$(.ID3.Artist)
        If chkTitle.Value = Checked Then sStr = sStr & Trim$(.ID3.Title)
        If chkGenre.Value = Checked Then sStr = sStr & Trim$(.ID3.Genre)
        If chkYear.Value = Checked Then sStr = sStr & Trim$(.ID3.Year)
        If chkVolume.Value = Checked Then sStr = sStr & Trim$(.VolumeName)
    End With
    
    
    m = InStr(1, sStr, Text4.Text, vbTextCompare)
    If m > 0 And Trim$(sStr) <> "" Then
    
        With IndexFile
        tmpItem$ = "Vol " & Trim(.VolumeName) & "> " & Trim$(.Filename)
        End With
        
        '// ADD TO LISTVIEW
        With ListView1
        
            .ListItems.Add matches + 1, , "Vol " & Trim$(IndexFile.VolumeName), 1, 1
            .ListItems(matches + 1).SubItems(1) = Trim$(IndexFile.Filename)
            .ListItems(matches + 1).SubItems(2) = Trim$(IndexFile.ID3.Album)
            .ListItems(matches + 1).SubItems(3) = Trim$(IndexFile.ID3.Artist)
            .ListItems(matches + 1).SubItems(4) = Trim$(IndexFile.ID3.Genre)
            .ListItems(matches + 1).SubItems(5) = Trim$(IndexFile.ID3.Title)
            .ListItems(matches + 1).SubItems(6) = Trim$(IndexFile.ID3.Year)
            .ListItems(matches + 1).SubItems(7) = Trim$(IndexFile.ID3.Comment)
            .ListItems(matches + 1).SubItems(8) = Trim$(Str$(vloop))
            
        End With
        
        matches = matches + 1
        Label4.Caption = matches & " matches found (" & vloop & " records)"
        Label4.Refresh
        
    End If
    
    ProgressBar1.Value = vloop
    
    If matches >= MaxVal Then Exit For
    If lblCancelFlag.Caption = "1" Then
        lblCancelFlag.Caption = "0"
        Exit For
    End If
    
    DoEvents
    Next
    
    Close ff
        
    ProgressBar1.Visible = False
    Label4.Caption = "Found " & matches & " matches out of " & vloop & " records."

    Label7.Caption = "Search: [" & Text4.Text & "]    " & matches & " Matches"
    ListView1.Sorted = True
    ListView1.Visible = True
    MousePointer = 0
    
    If ListView1.ListItems.Count > 0 Then
        Me.lblToText.Enabled = True
        Me.lblToHtml.Enabled = True
    Else
        Me.lblToText.Enabled = False
        Me.lblToHtml.Enabled = False
    End If
    

End Sub

Private Sub ExportToHtml()
    
    If lblToHtml.Enabled = False Then Exit Sub
    If ListView1.ListItems.Count < 1 Then Exit Sub

    HideMenu
    
    '// SHOW DETAILS OF SELECTED RECORD
    If ListView1.ListItems.Count = 0 Then Exit Sub
        
    MousePointer = 13
    gCancelProcess = False
    
    Me.picDeleteProgress.Visible = True
    Me.lblActivityTitle.Caption = "Exporting to HTML"
    Me.picDeleteProgress.Refresh
    Me.pbDelete.Value = 1
    Me.pbDelete.Max = 2
    Me.pbDelete.min = 0
    Me.pbDelete.Max = ListView1.ListItems.Count
    
    Dim hFile$, rFile$, HtmlString$
    hFile$ = ExeDir & "html-header-template.txt"
    rFile$ = ExeDir & "html-record-template.txt"
        
        
    '// PARSE TEMPLATES                                         //
    HtmlString = GetTextFromFile(hFile)
    HtmlString = Replace(HtmlString, "%search%", Label7.Caption)
    
    mbt$ = GetTextFromFile(rFile)
    
    For Y = 1 To ListView1.ListItems.Count
    
    DoEvents
    
    Me.lblCounter.Caption = "Exporting item " & Y & " of " & pbDelete.Max
    pbDelete.Value = Y
    Me.lblCounter.Refresh
    
        '// GET RECORD DATA                                         //
        ff = FreeFile
        Open CatalogFileName For Random As ff Len = Len(IndexFile)
        '// Record number is in column 9
        Get ff, ListView1.ListItems(Y).SubItems(8), IndexFile
        Close ff
        
        With IndexFile
            ThisRecord$ = Replace(mbt$, "%volume%", Trim$(.VolumeName))
            ThisRecord$ = Replace(ThisRecord$, "%filename%", Trim$(.Filename))
        End With
        
        With IndexFile.ID3
            ThisRecord$ = Replace(ThisRecord$, "%album%", Trim$(.Album))
            ThisRecord$ = Replace(ThisRecord$, "%artist%", Trim$(.Artist))
            ThisRecord$ = Replace(ThisRecord$, "%comments%", Trim$(.Comment))
            ThisRecord$ = Replace(ThisRecord$, "%genre%", Trim$(.Genre))
            ThisRecord$ = Replace(ThisRecord$, "%title%", Trim$(.Title))
            ThisRecord$ = Replace(ThisRecord$, "%year%", Trim$(.Year))
        End With
        
        With IndexFile.Mp3Info
            ThisRecord$ = Replace(ThisRecord$, "%bitrate%", Trim$(.BitRate))
            ThisRecord$ = Replace(ThisRecord$, "%copyright%", Trim$(.Copy))
            ThisRecord$ = Replace(ThisRecord$, "%crc%", Trim$(.CRC))
            ThisRecord$ = Replace(ThisRecord$, "%emphasis%", Trim$(.Emphasis))
            ThisRecord$ = Replace(ThisRecord$, "%frequency%", Trim$(.FreqChannel))
            ThisRecord$ = Replace(ThisRecord$, "%layer%", Trim$(.Layer))
            ThisRecord$ = Replace(ThisRecord$, "%length%", Trim$(.Length))
            ThisRecord$ = Replace(ThisRecord$, "%original%", Trim$(.Original))
            ThisRecord$ = Replace(ThisRecord$, "%size%", Trim$(.Size))
        End With
        
        HtmlString = HtmlString & ThisRecord$
        ThisRecord$ = ""
        If gCancelProcess = True Then Exit For
        
    Next
    

    
    '// WRITE TO FILE
    ff = FreeFile
    Open "c:\dp-tmp.htm" For Output As ff
    Print #ff, HtmlString
    Close ff
    
    Me.picDeleteProgress.Visible = False
    
    
    '// DISPLAY FILE IN BROWSER             //
    Dim Dummy As Long
    Dummy = ShellExecute(Me.hWnd, vbNullString, "c:\dp-tmp.htm", _
                         vbNullString, "c:\", 1)
    


    
    
    MousePointer = 0
    
End Sub

Private Sub ExportToText()

    If lblToText.Enabled = False Then Exit Sub
    If ListView1.ListItems.Count < 1 Then Exit Sub

    HideMenu
    
    '// SHOW DETAILS OF SELECTED RECORD
    If ListView1.ListItems.Count = 0 Then Exit Sub
        
    MousePointer = 11
    
        mbt$ = "Disk Pro" & vbCrLf _
             & Label7.Caption & vbCrLf _
             & "==================================================================" _
             & vbCrLf & vbCrLf
             
    For Y = 1 To ListView1.ListItems.Count
    
        mbt$ = mbt$ & "Volume: " & ListView1.ListItems(Y).Text & vbCrLf
        For X = 1 To 7
        mbt$ = mbt$ & ListView1.ColumnHeaders(X + 1).Text & ": " _
                & ListView1.ListItems(Y).SubItems(X) _
                & vbCrLf
        Next
        mbt$ = mbt$ & vbCrLf
        
    Next
    
    ff = FreeFile
    Open "c:\tmp.txt" For Output As ff
    Print #ff, mbt$
    Close ff
    MousePointer = 0
    
    
    Shell "notepad.exe c:\tmp.txt", vbNormalFocus
    
    
End Sub



Private Sub Command1_Click()

    ShowMenu 3
    
End Sub

Private Sub Command2_Click()

    HideMenu
    
End Sub

Private Sub Command3_Click()

    answer$ = MsgBox("Do you want to clear the keyword history?", _
                    vbYesNo + vbQuestion, "Are you sure?")
                    
    If answer <> vbYes Then Exit Sub
    
    '// CLEAR FILE  //
    kwfile$ = ExeDir & "\keywords.txt"
    ff = FreeFile
    Open kwfile For Random As ff
    Close ff
    Kill kwfile
    
    '// CLEAR LIST  //
    List1.Clear
    List1.SetFocus
    

End Sub

Private Sub Command4_Click()

    gCancelProcess = True
    
End Sub

Private Sub Form_Load()

    cmbMax.Text = Trim(Config.MaxList)
    ListView1.View = Trim(Config.ListviewStyle)
    
    chkFilename.Value = Config.sFilename
    chkAlbum.Value = Config.sAlbum
    chkArtist.Value = Config.sArtist
    chkTitle.Value = Config.sTitle
    chkGenre.Value = Config.sGenre
    chkYear.Value = Config.sYear
    chkVolume.Value = Config.sVolume
    chkClearList.Value = Config.iClearList
    
    With ListView1
        For X = 1 To 9
            .ColumnHeaders(X).Width = Config.chWidth(X)
        Next
    End With
    
    MousePointer = 11
    doFillList
    MousePointer = 0
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    HideMenu
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    Me.ProgressBar1.Width = ScaleWidth - 600
    ListView1.Width = ScaleWidth - 240
    ListView1.Height = ScaleHeight - ListView1.Top - 180
    Frame1.Move ListView1.Left, ListView1.Top, ListView1.Width, ListView1.Height
    cmdCancel.Left = ProgressBar1.Left + ProgressBar1.Width - cmdCancel.Width
    
    With imgLogo
        .Top = 15
        .Left = ScaleWidth - .Width - 15
    End With
    
End Sub

Private Sub HideMenu(Optional Index%)

    If Index Then
        
        Select Case Index
        
        Case 2
        Me.picOptMenu.Visible = False
        Me.picOptShadow.Visible = False
        
        Case 1
        Me.picExportMenu.Visible = False
        Me.picExportShadow.Visible = False
    
        End Select
        
    Else
        
        Me.picOptMenu.Visible = False
        Me.picOptShadow.Visible = False
        Me.picExportMenu.Visible = False
        Me.picExportShadow.Visible = False
        Me.picKeywordHistoryShadow.Visible = False
        Me.picSearchlist.Visible = False
    
    End If
    
        HideRightClickMenu
        
End Sub

Private Sub ShowMenu(Index%)

    Select Case Index
    
    Case 2
    Me.picOptMenu.Visible = True
    Me.picOptShadow.Visible = True
    
    Case 1
    Me.picExportMenu.Visible = True
    Me.picExportShadow.Visible = True
    
    Case 3
    Me.picKeywordHistoryShadow.Visible = True
    Me.picSearchlist.Visible = True
    
    End Select
    
End Sub


Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ShowMenu 2
    HideMenu 1
    
    
End Sub



Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ShowMenu 1
    HideMenu 2
    
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    HideMenu
    
End Sub

Private Sub Label23_Click()

End Sub

Private Sub Label24_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ReleaseCapture
    SendMessage Me.picDeleteProgress.hWnd, &HA1, 2, 0&
    
End Sub

Private Sub Label7_Change()

    Label5.Caption = Label7.Caption
    Caption = Label5.Caption
    
End Sub

Private Sub lblActivityTitle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    ReleaseCapture
    SendMessage Me.picDeleteProgress.hWnd, &HA1, 2, 0&
    
End Sub

Private Sub lblContextHL_Click()

    If lblPlay.Enabled = False Then Exit Sub
    
    If lblContextHL.Top = lblPlay.Top - 85 Then
        lblPlay_Click
    ElseIf lblContextHL.Top = lblDelete.Top - 85 Then
        lblDelete_Click
    End If
    
        
End Sub

Private Sub lblDelete_Click()

    DeleteEntries
    
End Sub

Private Sub lblDelete_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.lblContextHL.Top = lblDelete.Top - 85
    Me.lblContextHL.Visible = True
    
End Sub

Private Sub lblExportHL_Click()

    If lblToHtml.Enabled = False Then Exit Sub
    
    If lblExportHL.Top = lblToText.Top - 85 Then
        lblToText_Click
    ElseIf lblExportHL.Top = lblToHtml.Top - 85 Then
        lblToHtml_Click
    End If
    
End Sub

Private Sub lblPlay_Click()
    
    ListView1_DblClick
    HideRightClickMenu
    
End Sub

Private Sub lblPlay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.lblContextHL.Top = lblPlay.Top - 85
    Me.lblContextHL.Visible = True
    
End Sub

Private Sub lblToHtml_Click()

    ExportToHtml
    
End Sub

Private Sub lblToHtml_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.lblExportHL.Top = lblToHtml.Top - 85
    Me.lblExportHL.Visible = True
   
End Sub

Private Sub lblToText_Click()

    ExportToText
    
End Sub

Private Sub lblToText_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.lblExportHL.Top = lblToText.Top - 85
    Me.lblExportHL.Visible = True
    
End Sub

Private Sub List1_Click()

    On Error Resume Next
    Me.Text4.Text = List1.List(List1.ListIndex)
    HideMenu
    cmdSearch_Click
    
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    ListView1.SortKey = ColumnHeader.Index - 1
    
    For X = 1 To 9
    ListView1.ColumnHeaders(X).Text = Trim$(Replace(ListView1.ColumnHeaders(X).Text, "*", ""))
    Next
    
    ListView1.ColumnHeaders(ColumnHeader.Index).Text = "* " & ListView1.ColumnHeaders(ColumnHeader.Index).Text
    
End Sub



Private Sub ListView1_DblClick()

    '// SHOW DETAILS OF SELECTED RECORD
    If ListView1.ListItems.Count = 0 Then Exit Sub
       
    MousePointer = 13
    
    mbt$ = "Volume: " & ListView1.SelectedItem.Text & vbCrLf
    For X = 1 To 8
    mbt$ = mbt$ & ListView1.ColumnHeaders(X + 1).Text & ": " _
                & ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(X) _
                & vbCrLf
    Next
    
    ff = FreeFile
    Open CatalogFileName For Random As ff Len = Len(IndexFile)
    '// Record number is in column 9
    Get ff, ListView1.SelectedItem.SubItems(8), IndexFile
    Close ff
        
    With IndexFile.Mp3Info
    mbt$ = mbt$ & "Size: " & .Size & vbCrLf _
                & "Length: " & .Length & vbCrLf _
                & "Layer: " & .Layer & vbCrLf _
                & "BitRate: " & .BitRate & vbCrLf _
                & "Frequency Channel: " & .FreqChannel & vbCrLf _
                & "CRC: " & .CRC & vbCrLf _
                & "Copyright: " & .Copy & vbCrLf _
                & "Emphasis: " & .Emphasis & vbCrLf _
                & "Original: " & .Original & vbCrLf
    End With
    
    PlayMP3 Trim$(Config.iDefaultPlayerDrive) _
          & ListView1.ListItems(ListView1.SelectedItem.Index).SubItems(1), _
          mbt$, ListView1.SelectedItem.Text
          
    MousePointer = 0
    
    
End Sub

Private Sub ListView1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    HideMenu
    
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Select Case Button
    
    Case 1      'LEFT
    HideRightClickMenu
    Me.lblPlay.Enabled = False
    Me.lblDelete.Enabled = False
    
    Case 2
    
    Me.picRightClickMenu.Left = X + ListView1.Left
    Me.picRightClickMenu.Top = Y + ListView1.Top
    Me.picRightClickMenuShadow.Left = picRightClickMenu.Left + 30
    Me.picRightClickMenuShadow.Top = picRightClickMenu.Top + 30
    ShowRightClickMenu
    
    If ListView1.ListItems.Count > 0 Then
        If ListView1.SelectedItem.Text <> "" Then
            Me.lblPlay.Enabled = True
            Me.lblDelete.Enabled = True
        End If
    End If
    
    '// MAKE SURE CONTEXT MENU IS IN VIEW
    Do Until (picRightClickMenu.Left + picRightClickMenu.Width + 30) < (ListView1.Left + ListView1.Width)
        picRightClickMenu.Left = picRightClickMenu.Left - 15
        picRightClickMenuShadow.Left = picRightClickMenu.Left + 30
    Loop
    
    Do Until (picRightClickMenu.Top + picRightClickMenu.Height + 30) < (ListView1.Top + ListView1.Height)
        picRightClickMenu.Top = picRightClickMenu.Top - 15
        picRightClickMenuShadow.Top = picRightClickMenu.Top + 30
    Loop
    
    End Select
    
End Sub

Private Sub picExportMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    lblExportHL.Visible = False
    
End Sub

Private Sub picRightClickMenu_LostFocus()

    HideRightClickMenu
    
End Sub

Private Sub picRightClickMenu_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Me.lblContextHL.Visible = False
    
End Sub

Private Sub Text4_Change()

    cmdSearch.Default = True
    HideMenu
    
End Sub

Private Sub Timer1_Timer()

    Static OldSum
    Dim NewSum As Integer
    
    For X = 1 To 9
    NewSum% = NewSum% + ListView1.ColumnHeaders(X).Width
    Next
    
    If OldSum <> NewSum Then
        '// SAVE NEW COLUMN SIZES
        For X = 1 To 9
        Config.chWidth(X) = ListView1.ColumnHeaders(X).Width
        Next
        SaveConfig
    End If
    
    OldSum = NewSum
    
End Sub

Private Sub doAddKeyword()

    keyword$ = Text4.Text
    
    If List1.ListCount < 1 Then GoTo doAdd
    
    For X = 0 To List1.ListCount
        If LCase(List1.List(X)) = LCase$(keyword) Then
            Flag% = 1
            Exit For
        End If
    Next
    
    If Flag = 1 Then Exit Sub
    
    
    
doAdd:
    kwfile$ = ExeDir & "\keywords.txt"
    
    ff = FreeFile
    Open kwfile For Random As ff
    Close ff
    
    Open kwfile For Append As ff
    Print #ff, keyword
    Close ff
    
    Me.List1.AddItem keyword
    
End Sub

Private Sub doFillList()

    List1.Clear
    kwfile$ = ExeDir & "\keywords.txt"
    
    ff = FreeFile
    Open kwfile For Random As ff
    Close ff
    
    Open kwfile For Input As ff
    Do Until EOF(ff)
        Line Input #ff, a$
        List1.AddItem a$
    Loop
    Close ff
    
End Sub
