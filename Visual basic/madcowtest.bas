Attribute VB_Name = "sndMain"
'// Plays Midi, Mpeg, Mp3 Files (.mid .mpeg .mp3)
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'// Plays Wave Files (.wav)
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    Public Const SND_SYNC = 0 ' Freeze program until sound is done
    Public Const SND_ASYNC = 1 ' Continue program while sound plays
    Public Const SND_NODEFAULT = 2 ' Don't beep if sound not found
    Public Const SND_LOOP = 8
    Public Const SND_NOSTOP = 16 ' The function returns false if another sound is playing

