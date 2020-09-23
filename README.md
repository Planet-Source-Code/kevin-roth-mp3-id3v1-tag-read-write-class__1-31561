<div align="center">

## MP3 ID3v1 Tag Read/Write Class


</div>

### Description

Use this class to read and/or write to the ID3 tag of an MP3.
 
### More Info
 
Must pass full file paths to the gettag and writetag subs


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Kevin Roth](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/kevin-roth.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Sound/MP3](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/sound-mp3__1-45.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/kevin-roth-mp3-id3v1-tag-read-write-class__1-31561/archive/master.zip)





### Source Code

```
Public ValidTag As String
Public Title As String
Public Artist As String
Public Year As String
Public Album As String
Public Comment As String
Public Genre As Byte
Private Type ID3v1
  ValidTag As String * 3
  Title As String * 30
  Artist As String * 30
  Album As String * 30
  Year As String * 4
  Comment As String * 30
  Genre As Byte
End Type
Public Sub getTag(MP3 As String)
  Dim ID3 As ID3v1
  Open MP3 For Binary As #1
  Get #1, FileLen(MP3) - 127, ID3
  Close #1
  With ID3
   ValidTag = .ValidTag
   Title = .Title
   Artist = .Artist
   Album = .Album
   Comment = .Comment
   Year = .Year
   Genre = .Genre
  End With
End Sub
Public Sub writeTag(MP3 As String)
  Dim ID3 As ID3v1
  With ID3
   .ValidTag = "TAG"
   .Title = Title
   .Artist = Artist
   .Album = Album
   .Comment = Comment
   .Year = Year
   .Genre = Genre
  End With
  On Error GoTo ErrMsg:
  Open MP3 For Binary As 1
  If ID3.ValidTag <> "TAG" Then
    Seek 1, LOF(1) + 1
  Else
    Seek 1, LOF(1) - 127
  End If
  Put 1, FileLen(MP3) - 127, ID3
  Close 1
  Exit Sub
ErrMsg:
  MsgBox ("File '" & MP3 & "' is marked as read-only or the file is in use." & vbCr & "Please correct and try again.")
End Sub
```

