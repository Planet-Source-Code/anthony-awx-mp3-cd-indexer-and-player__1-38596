VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "idtag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private mvarFilename As String

Private Type Info
    
    sTitle As String
    sArtist As String
    sAlbum As String
    sComment As String
    sYear As String
    sGenre As String

End Type

Private Mp3Info As Info

Public Property Get Filename() As String
    
    Filename = mvarFilename

End Property

Private Function IsValidFile(ByVal sFilename) As Boolean
    
    Dim bOk As Boolean
    
    Do Until InStr(1, sFilename, "\\") = 0
        sFilename = Replace(sFilename, "\\", "\")
    Loop
    
    ' make sure file exists                     '
    bOk = CBool(Dir(sFilename, vbHidden) <> "")
    
    Dim aExtensions, ext
    aExtensions = Array(".mp3", ".mp2", ".mp1")
    Dim bOkayExtension As Boolean
    bOkayExtension = False

    If bOk Then


        For Each ext In aExtensions


            If InStr(1, sFilename, ext, vbTextCompare) > 0 Then
                bOkayExtension = True
            End If
        Next 'ext
    End If
    
    IsValidFile = bOk And bOkayExtension

End Function


Public Property Let Filename(ByVal sPassFilename As String)
    
    Do Until InStr(1, sPassFilename, "\\") = 0
        sPassFilename = Replace(sPassFilename, "\\", "\")
    Loop
    
    Dim iFreefile As Integer
    Dim lFilePos As Long
    Dim sData As String * 128
    
    Dim sGenre() As String
    
    ' Genre '
    Const sGenreMatrix As String = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
    "Hip-Hop|Jazz|Metal|New Age|Oldies|Other|Pop|R&B|Rap|Reggae|Rock|Techno|" + _
    "Industrial|Alternative|Ska|Death Metal|Pranks|Soundtrack|Euro-Techno|" + _
    "Ambient|Trip Hop|Vocal|Jazz+Funk|Fusion|Trance|Classical|Instrumental|Acid|" + _
    "House|Game|Sound Clip|Gospel|Noise|Alt. Rock|Bass|Soul|Punk|Space|Meditative|" + _
    "Instrumental Pop|Instrumental Rock|Ethnic|Gothic|Darkwave|Techno-Industrial|Electronic|" + _
    "Pop-Folk|Eurodance|Dream|Southern Rock|Comedy|Cult|Gangsta Rap|Top 40|Christian Rap|" + _
    "Pop/Punk|Jungle|Native American|Cabaret|New Wave|Phychedelic|Rave|Showtunes|Trailer|" + _
    "Lo-Fi|Tribal|Acid Punk|Acid Jazz|Polka|Retro|Musical|Rock & Roll|Hard Rock|Folk|" + _
    "Folk/Rock|National Folk|Swing|Fast-Fusion|Bebob|Latin|Revival|Celtic|Blue Grass|" + _
    "Avantegarde|Gothic Rock|Progressive Rock|Psychedelic Rock|Symphonic Rock|Slow Rock|" + _
    "Big Band|Chorus|Easy Listening|Acoustic|Humour|Speech|Chanson|Opera|Chamber Music|" + _
    "Sonata|Symphony|Booty Bass|Primus|Porn Groove|Satire|Slow Jam|Club|Tango|Samba|Folklore|" + _
    "Ballad|power Ballad|Rhythmic Soul|Freestyle|Duet|Punk Rock|Drum Solo|A Capella|Euro-House|" + _
    "Dance Hall|Goa|Drum & Bass|Club-House|Hardcore|Terror|indie|Brit Pop|Negerpunk|Polsk Punk|" + _
    "Beat|Christian Gangsta Rap|Heavy Metal|Black Metal|Crossover|Comteporary Christian|" + _
    "Christian Rock|Merengue|Salsa|Trash Metal|Anime|JPop|Synth Pop"
    
    sGenre = Split(sGenreMatrix, "|")
    mvarFilename = sPassFilename
    
    If Not IsValidFile(sPassFilename) Then
        Exit Property
    End If
    
    Mp3Info.sTitle = ""
    Mp3Info.sArtist = ""
    Mp3Info.sAlbum = ""
    Mp3Info.sYear = ""
    Mp3Info.sComment = ""
    
    iFreefile = FreeFile
    lFilePos = FileLen(mvarFilename) - 127


    If lFilePos > 0 Then
        Open mvarFilename For Binary As #iFreefile
        Get #iFreefile, lFilePos, sData
        Close #iFreefile
    End If
    
    ' Populate                                      '
    If Left(sData, 3) = "TAG" Then
        Mp3Info.sTitle = Mid(sData, 4, 30)
        Mp3Info.sArtist = Mid(sData, 34, 30)
        Mp3Info.sAlbum = Mid(sData, 64, 30)
        Mp3Info.sYear = Mid(sData, 94, 4)
        Mp3Info.sComment = Mid(sData, 98, 30)
        Dim lGenre
        lGenre = Asc(Mid(sData, 128, 1))

        If lGenre <= UBound(sGenre) Then
            Mp3Info.sGenre = sGenre(lGenre)
        Else
            Mp3Info.sGenre = ""
        End If
    
    Else
        Mp3Info.sArtist = ""
        Mp3Info.sAlbum = ""
        Mp3Info.sComment = ""
        Mp3Info.sGenre = ""
        Mp3Info.sTitle = "NA"
        Mp3Info.sYear = ""
    End If

End Property

Private Function GetInfo(ByVal sFilename) As Info
    
    Dim i As Info
    GetInfo = i
    Dim s
    s = sFilename

    If InStrRev(s, "\") > 0 Then
        s = Mid(s, InStrRev(s, "\") + 1)
    End If
    
    s = Left(s, InStrRev(s, ".", , vbTextCompare) - 1)
    s = Replace(Trim(s), " ", " ")
    s = Trim(s)
    
    If CountItems(s, " ") < 1 Then
        i.sTitle = Replace(s, "_", " ")
        GetInfo = i
        Exit Function
    End If
    
    s = Trim(Replace(s, "_", " "))

    If Left(s, 1) = "(" And CountItems(s, "-") < 3 Then
        i.sArtist = Mid(s, 2, InStr(s, ")") - 2)
        s = Trim(Mid(s, InStr(s, ")") + 1))
        If Left(s, 1) = "-" Then
            i.sTitle = Trim(Mid(s, 2))
        
        Else
            If InStr(s, "-") > 0 Then
                i.sAlbum = Mid(s, InStr(s, "-") + 1)
                i.sTitle = Left(s, InStr(s, "-") - 1)
            Else
                i.sTitle = Trim(s)
            End If
        
        End If
    
    Else
        Dim aThings
        Dim l
        aThings = Split(s, "- ")

        For l = 0 To UBound(aThings)
            If Not IsNumeric(aThings(l)) Then
                If i.sArtist = "" Then
                    i.sArtist = aThings(l)
                
                Else
                    If IsNumeric(aThings(l - 1)) Then

                        If i.sTitle = "" Then
                            i.sTitle = aThings(l)
                        End If
                    
                    ElseIf i.sAlbum = "" Then
                        i.sAlbum = aThings(l)
                    
                    End If
                
                End If
            
            End If
        
        Next
        
    End If
    
    i.sArtist = Replace(Replace(i.sArtist, "(", ""), ")", "")
    
    If Left(s, 1) <> "(" And i.sTitle = "" And (InStr(sFilename, "\") <> InStrRev(sFilename, "\")) Then
        ' CALL THIS ROUTINE AGAIN RECURSIVELY   '
        GetInfo = GetInfo(FixDir(sFilename))
    
    Else
        GetInfo = i
    
    End If

End Function

Private Function CountItems(s, sToCount)
    
    Dim a
    a = Split(s, sToCount)

    If UBound(a) = -1 Then
        CountItems = 0
    
    Else
        CountItems = UBound(a) - LBound(a)
    
    End If

End Function

Private Function FixDir(sFullpath)
    
    Dim s1, s2
    s1 = Trim(Left(sFullpath, InStrRev(sFullpath, "\") - 1))
    s2 = Trim(Mid(sFullpath, InStrRev(sFullpath, "\") + 1))
    FixDir = s1 & " - " & s2

End Function


Public Property Get Title() As String
    
    Title = Trim(Mp3Info.sTitle)

End Property

Public Property Get Artist() As String
    
    Artist = Trim(Mp3Info.sArtist)

End Property

Public Property Get Genre() As String
    
    Genre = Trim(Mp3Info.sGenre)

End Property


Public Property Get Album() As String
    
    Album = Trim(Mp3Info.sAlbum)

End Property

Public Property Get Year() As String
    
    Year = Trim(Mp3Info.sYear)

End Property

Public Property Get Comment() As String
    
    Comment = Trim(Mp3Info.sComment)

End Property


