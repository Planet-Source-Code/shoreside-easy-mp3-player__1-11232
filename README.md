<div align="center">

## Easy MP3 Player


</div>

### Description

One of the easiest MP3 Player's! Don't understand other MP3 complex code...this is for you!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[ShoreSide](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/shoreside.md)
**Level**          |Beginner
**User Rating**    |4.5 (36 globes from 8 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Sound/MP3](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/sound-mp3__1-45.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/shoreside-easy-mp3-player__1-11232/archive/master.zip)





### Source Code

```
' Simple MP3 Player
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Dim isPlaying As Boolean
Dim Mp3File As String
Private Sub Command1_Click(Index As Integer)
  Mp3File = Chr$(34) + Trim(Text1.Text) + Chr$(34)
  Select Case Index
   Case 0
    ' Start Playing
    mciSendString "open " + Mp3File, 0&, 0&, 0&
    mciSendString "play " + Mp3File, "", 0&, 0&
    isPlaying = True
   Case 1
    ' Stop Playing
    mciSendString "close " + Mp3File, 0&, 0&, 0&
    isPlaying = False
  End Select
End Sub
Private Sub Command2_Click()
  Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If isPlaying = True Then
   ' Stop Playing if we are playing before we exit!
   mciSendString "close " + Mp3File, 0&, 0&, 0&
  End If
End Sub
```

