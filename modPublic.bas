Attribute VB_Name = "modPublic"
Type ID3
    Album As String * 30
    Artist As String * 30
    Comment As String * 30
    Genre As String * 30
    Title As String * 30
    Year As String * 4
End Type

Type Mp3Inf
    Size As String * 20
    Length As String * 20
    Layer As String * 15
    BitRate As String * 15
    FreqChannel As String * 15
    CRC As String * 15
    Copy As String * 15
    Emphasis As String * 15
    Original As String * 15
End Type
    
Type OldIndex
    VolumeName As String * 20
    Filename As String * 250
End Type
    
Type IndexFile
    VolumeName As String * 20
    Filename As String * 250
    ID3 As ID3
    Mp3Info As Mp3Inf
End Type

Type cfg
    Initialized As Boolean
    wLeft As Integer
    wTop As Integer
    wWidth As Integer
    wHeight As Integer
    wState As Integer
    DBFileName As String * 300
    MaxList As Integer
    chWidth(9) As Integer
    ListviewStyle As String * 20
    sFilename As Integer
    sAlbum As Integer
    sArtist As Integer
    sTitle As Integer
    sGenre As Integer
    sYear As Integer
    iAlign As Integer
    iInterface As Integer
    iSearchWindows As Integer
    iSearchAtStartup As Integer
    iDefaultPlayerDrive As String * 6
    sVolume As Integer
    iClearList As Integer
    Future As String * 986
End Type

Declare Function ShellExecute _
                 Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hWnd As Long, _
                 ByVal lpOperation As String, _
                 ByVal lpFile As String, _
                 ByVal lpParameters As String, _
                 ByVal lpDirectory As String, _
                 ByVal nShowCmd As Long) As Long
             
Declare Function ReleaseCapture Lib "user32" () As Long

Public Declare Function SendMessage _
    Lib "user32" Alias "SendMessageA" _
    (ByVal hWnd As Long, ByVal wMsg As Long, _
     ByVal wParam As Integer, ByVal lParam As Any) As Long
     
Declare Function SetWindowPos Lib "user32" _
        (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
         ByVal X As Long, ByVal Y As Long, ByVal cx As Long, _
         ByVal cy As Long, ByVal wFlags As Long) As Long
         
Public IndexFile As IndexFile
Public CatalogFileName As String
Public cfgFile As String, Config As cfg
Public idtag As idtag, ExeDir As String
Public gCancelProcess As Boolean


Public Sub SaveConfig()

    ff = FreeFile
    Open cfgFile For Random As ff Len = Len(Config)
    Put ff, 1, Config
    Close ff
        
End Sub

Public Sub GetConfig()

    ff = FreeFile
    Open cfgFile For Random As ff Len = Len(Config)
    Get ff, 1, Config
    Close ff
        
End Sub

