VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsID3v1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'读取MP3中信息
Private Type tID3v1             'This type is standard for ID3v1 tags
  Title       As String * 30    '30 bytes Title
  Artist      As String * 30    '30 bytes Artist
  Album       As String * 30    '30 bytes Album
  Year        As String * 4     '4 bytes Year
  Comments    As String * 28    '28 bytes Comments
  IsTrack     As Byte           '1 byte Istrack / +1 byte comments
  Tracknumber As Byte           '1 byte Tracknumber / +1 byte comments
  Genre       As Byte           '1 byte Genre
End Type

Private Const sGenreMatrix = "Blues|Classic Rock|Country|Dance|Disco|Funk|Grunge|" + _
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

Private Genre() As String  'this array holds genres
Private ID3V1Info As tID3v1
Private strFilename As String

Public Property Get filename() As String
  filename = strFilename
End Property

Public Property Let filename(ByVal sFileName As String)
  strFilename = sFileName
End Property

Private Sub Class_Initialize()
  Genre = Split(sGenreMatrix, "|")
End Sub

Public Function ReadTag() As Boolean
  On Error Resume Next
  Dim sTag As String * 3   'used to grab tag
  Dim iFile As Integer

  On Error GoTo ErrH
  
  iFile = FreeFile
  ReadTag = True
  
  Open strFilename For Binary As iFile
  Get iFile, FileLen(strFilename) - 127, sTag        'Looks after tag 128 bytes from the end of the file
  If sTag = "TAG" Then                            '"TAG" is put at position filesize-127 to show that this file indeed contains an ID3v1 tag
    Get iFile, FileLen(strFilename) - 124, ID3V1Info 'if the file has a tag, we put it into our earlier declared variable ID3v1Info
  Else
    ID3V1Info.Title = ""                          'if the "TAG" wasnt at position filesize-127
    ID3V1Info.Artist = ""
    ID3V1Info.Album = ""
    ID3V1Info.Year = ""
    ID3V1Info.Comments = ""
    ID3V1Info.IsTrack = 0
    ID3V1Info.Tracknumber = 0
    ID3V1Info.Genre = 255
    ReadTag = False
  End If
  Close iFile

  Exit Function
ErrH:
  ReadTag = False
  Close iFile
End Function

Public Property Get tagTitle() As String
  tagTitle = sNT(Trim(ID3V1Info.Title))
End Property

Public Property Let tagTitle(ByVal sTitle As String)
  ID3V1Info.Title = sTitle
End Property

Public Property Get tagArtist() As String
  tagArtist = sNT(Trim(ID3V1Info.Artist))
End Property

Public Property Let tagArtist(ByVal sArtist As String)
  ID3V1Info.Artist = sArtist
End Property

Public Property Get tagAlbum() As String
  tagAlbum = sNT(Trim(ID3V1Info.Album))
End Property

Public Property Let tagAlbum(ByVal sAlbum As String)
  ID3V1Info.Album = sAlbum
End Property

Public Property Get tagStrGenre() As String
  If ID3V1Info.Genre <= UBound(Genre) Then
    tagStrGenre = Genre(ID3V1Info.Genre)
  End If
End Property

Public Property Get tagGenre() As Byte
  tagGenre = ID3V1Info.Genre
End Property

Public Property Let tagGenre(ByVal bGenre As Byte)
  ID3V1Info.Genre = bGenre
End Property

Public Property Get tagComments() As String
  If ID3V1Info.IsTrack <> 0 Then
    tagComments = sNT(Trim(ID3V1Info.Comments & Chr(ID3V1Info.IsTrack) & Chr(ID3V1Info.Tracknumber)))
  Else
    tagComments = sNT(Trim(ID3V1Info.Comments))
  End If
End Property

Public Property Let tagComments(ByVal sComments As String)
  'NOTE: if you do not want to write over tracknumbers, make sure the new
  'comment isn't longer than 28 chars.
  If Len(sComments) <= 28 Then
    ID3V1Info.Comments = sComments
  Else
    ID3V1Info.Comments = sComments
    ID3V1Info.IsTrack = Asc(Mid(sComments, 29, 1))
    ID3V1Info.Tracknumber = Asc(Mid(sComments, 30, 1))
  End If
End Property

Public Property Get tagTrack() As Byte
  If ID3V1Info.IsTrack = 0 Then
    tagTrack = ID3V1Info.Tracknumber
  Else
    tagTrack = 0
  End If
End Property

Public Property Let tagTrack(ByVal bTrack As Byte)
  ID3V1Info.IsTrack = 0
  ID3V1Info.Tracknumber = bTrack
End Property

Public Property Get tagYear() As String
  tagYear = sNT(Trim(ID3V1Info.Year))
End Property

Public Property Let tagYear(ByVal sYear As String)
  ID3V1Info.Year = sYear
End Property

Public Function ReturnGenre(ByVal bGenreNum As Byte) As String
  'This function returns the genre in an string
  If bGenreNum <= UBound(Genre) And bGenreNum >= LBound(Genre) Then
    ReturnGenre = Genre(bGenreNum)
  End If
End Function

Public Function SaveTag() As Boolean
  Dim sTag As String * 3
  Dim iFile As Integer
  
  On Error GoTo ErrH
  SaveTag = True
  iFile = FreeFile

  Open strFilename For Binary As iFile
  Get iFile, FileLen(strFilename) - 127, sTag    'Id3 tags are at the end of the mp3 file(and as the type shows it is 128 bytes)
  
  If sTag = "TAG" Then                             '"TAG" is put at position filesize-127 to show that this file indeed contains an Id3
    Put iFile, FileLen(strFilename) - 124, ID3V1Info   'if the file has a tag, we put our new information in the file
  Else
    Put iFile, FileLen(strFilename) - 127, "TAG"     'else we put the "TAG" there first,
    Close iFile
    Call SaveTag                                      'then we call this function again so we fill the info this time
  End If
  Close iFile
  
  Exit Function
ErrH:
  SaveTag = False
End Function

Public Function RemoveTag() As Boolean
  'This will remove the ID3v1 tag from an mp3 file
  Dim filedata() As Byte  'The mp3 will be read into this array
  Dim sTag As String * 3  'We use this variable to make sure the file has an ID3TAG
  Dim iFile As Integer
  Dim iFile2 As Integer

  On Error GoTo ErrH
  RemoveTag = True
  iFile = FreeFile 'Get free file number

  Open strFilename For Binary As iFile 'Open file
  Get iFile, FileLen(strFilename) - 127, sTag 'Look for the tag
    
  If sTag = "TAG" Then 'If there is an tag,
    ReDim filedata(FileLen(strFilename) - 129) 'redim array to fit mp3 file without tag
    
    iFile2 = FreeFile 'get another free file number
    Open strFilename & ".temp" For Binary As iFile2 'Open an temporary file
    Get iFile, 1, filedata                          'Reads whole mp3 without tag
    Put iFile2, 1, filedata                         'Writes to temporary file
    Close iFile2                                    'Closes files
    Close iFile
      
    Kill filename                                   'deletes old mp3
      
    Name filename & ".temp" As filename             'renames new file to the old ones name
    
  Else 'A tag was not found... abort!
    Close iFile
    RemoveTag = True
  End If

  Exit Function
ErrH:
  RemoveTag = False
End Function

Public Function HasTag() As Boolean
  'This function quickly checks if the mp3 has an tag
  'the difference between this and ReadTag is that if there is an tag,
  'this function will not read it
  Dim sTag As String * 3   'used to grab tag
  Dim iFile As Integer

  On Error GoTo ErrH
  
  iFile = FreeFile
  HasTag = True
  
  Open strFilename For Binary As iFile
  Get iFile, FileLen(strFilename) - 127, sTag
  If sTag <> "TAG" Then
    HasTag = False
  End If
  Close iFile

  Exit Function
ErrH:
  HasTag = False
  Close iFile
End Function


