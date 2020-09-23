<div align="center">

## Record a WAV \(\*\.wav\) file


</div>

### Description

Finally, the RECORD source code for WAV files is here. Very easy, yet getting this (and all MCI commands) from Microsoft is like pulling teeth. We need simple code, so I brought this to you. Enjoy. I believe that I will start a website very soon, that will contain the entire listing and usage of the MCI Commands for the ordinary programmer like ourselves.
 
### More Info
 
Create a form (Form1). Add 3 command buttons to the form (Command1 Command2 Command3).

Creates a WAV file


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Patrick K\. Bigley](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/patrick-k-bigley.md)
**Level**          |Unknown
**User Rating**    |5.0 (25 globes from 5 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/patrick-k-bigley-record-a-wav-wav-file__1-1938/archive/master.zip)

### API Declarations

```
'Included in the code below. Nothing here...(don't copy)
```


### Source Code

```
Private Declare Function mciSendString Lib "winmm.dll" Alias _
     "mciSendStringA" (ByVal lpstrCommand As String, ByVal _
     lpstrReturnString As Any, ByVal uReturnLength As Long, ByVal _
     hwndCallback As Long) As Long
Private Sub Command1_Click()
i = mciSendString("open new type waveaudio alias capture", 0&, 0, 0)
i = mciSendString("set capture bitspersample 8", 0&, 0, 0)
i = mciSendString("set capture samplespersec 11025", 0&, 0, 0)
i = mciSendString("set capture channels 1", 0&, 0, 0)
i = mciSendString("record capture", 0&, 0, 0)
'bitspersample can be:
'  8
'  16
'
'samplespersec can be:
'  11025
'  22050
'  44100
'
'channels can be:
' 1 = mono
' 2 = stereo
End Sub
Private Sub Command2_Click()
  i = mciSendString("stop capture", 0&, 0, 0)
  i = mciSendString("save capture c:\NewWave.wav", 0&, 0, 0)
'  i = mciSendString("close capture", 0&, 0, 0)
End Sub
Private Sub Command3_Click()
i = mciSendString("play capture from 0", 0&, 0, 0)
End Sub
Private Sub Form_Load()
Me.Caption = "WAVE RECORDER"
Command1.Caption = "Record"
Command2.Caption = "Stop"
Command3.Caption = "Play"
End Sub
Private Sub Form_Unload(Cancel As Integer)
i = mciSendString("close capture", 0&, 0, 0)
End Sub
```

