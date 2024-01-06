VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, _
    ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
    
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal _
    lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long

Private Const SND_ASYNC = &H1            '  play asynchronously
Private Const SND_MEMORY = &H4           '  lpszSoundName points to a memory file

Private ret As Long, explosion As String
Public pause As Boolean

Sub Initialize()
Dim bgSound As String
    bgSound = ThisWorkbook.path & "\assets\audio\breakout.mp3"
    ret = mciSendString("open """ & bgSound & """ Type MPEGVideo Alias Track1", 0&, 0, 0)
    explosion = LoadSounds("explosion")
End Sub

Private Function LoadSounds(fName As String) As String
Dim f As Long, str(1) As String
    str(0) = ThisWorkbook.path & "\assets\audio\" & fName & ".wav"
    f = FreeFile
    str(1) = Space(FileLen(str(0)))
    Open str(0) For Binary As #f
    Get #f, , str(1)
    Close #f
    LoadSounds = Trim(str(1))
End Function

Sub PauseBG()
    If pause Then _
    ret = mciSendString("pause Track1", 0&, 0, 0)
End Sub

Sub ResumeBG()
    If Not pause Then _
    ret = mciSendString("resume Track1", 0&, 0, 0)
End Sub

Sub Terminate()
    ret = mciSendString("stop all", 0, 0, 0)
    ret = mciSendString("close all", 0, 0, 0)
    ret = 0
End Sub

Sub PlayBGSound()
    On Error GoTo myErr
    DoEvents
    ret = mciSendString("play Track1 repeat", 0&, 0, 0)
myErr:
    If Err.Number = -2147418105 Then Exit Sub
End Sub

Sub PlayExplosion()
    Call PlaySound(explosion, 0, SND_ASYNC Or SND_MEMORY)
End Sub

