VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MIDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' MIDI出力デバイス数を取得する
Private Declare PtrSafe Function midiOutGetNumDevs Lib "winmm" () As Integer
' MIDIデバイスを開く
Private Declare PtrSafe Function midiOutOpen Lib "winmm.dll" (lphMidiOut As LongPtr, ByVal uDeviceID As Long, ByVal dwCallback As LongPtr, ByVal dwInstance As LongPtr, ByVal dwflags As Long) As Long
' MIDIデバイスから音をだす
Private Declare PtrSafe Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As LongPtr, ByVal dwMsg As Long) As Long
' MIDIデバイスを閉じる
Private Declare PtrSafe Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As LongPtr) As Long

Private Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
 
 
 
Private Handle As LongPtr   'MIDIデバイスのハンドル

Private CSettings As New Settings


Private Sub Class_Initialize()
    Dim ret As Long
    
    'MIDIデバイス数を取得する
    ret = midiOutGetNumDevs
    
    If ret = 0 Then
        MsgBox ("MIDI音源がないのでリアルタイム演奏はご利用できません")
    Else
        'MIDIデバイスを開く
        ret = midiOutOpen(Handle, -1, 0, 0, 0)
    End If
End Sub

Private Sub Play_Sound(ByVal SoundNum As Integer)
    Dim ret As Long
    
    Dim Msg As Long
    Msg = &H7F0C90 + (88 - SoundNum) * 256
    ret = midiOutShortMsg(Handle, Msg)
    
End Sub

Private Sub Stop_Sound(ByVal SoundNum As Integer)
    Dim ret As Long
    
    Dim Msg As Long
    Msg = &HC90 + (88 - SoundNum) * 256
    ret = midiOutShortMsg(Handle, Msg)
    
End Sub




'リアルタイム演奏を開始
Public Sub Start(ByVal StartColumn As Long)
    Is_Playing = True
    
    Dim ret As Long
    
    Dim t1 As Double, t2 As Double, Sleep_Time As Long
    
    
    Dim BPM As Integer: BPM = CSettings.Get_Tempo
    Dim ScoreType As String: CSettings.Get_ScoreLength
    Dim BasicOneCellSeconds As Double: BasicOneCellSeconds = CSettings.Get_OneCellSeconds(BPM, ScoreType)
    Dim OneCellSeconds As Double
    Dim NoteColor As Long: NoteColor = CSettings.Get_NoteColor()
    
    Dim Counter As Long: Counter = StartColumn
    If PIANOROLL_END_COLUMN = -1 Then
        MsgBox ("終わりを設定してください。")
        Call End_Play
        Exit Sub
    End If
    
    Do While (Is_Playing = True And Counter <= PIANOROLL_END_COLUMN)
        t1 = Timer
        With ThisWorkbook.Worksheets("ピアノロール")
            '連符があるか調べ、1セル当たりの秒数を設定
            If .Cells(NOTE_LENGTH_ROW, Counter).Value <> "" Then
                OneCellSeconds = 2 * BasicOneCellSeconds / CInt(.Cells(NOTE_LENGTH_ROW, Counter).Value)
            Else
                OneCellSeconds = BasicOneCellSeconds
            End If
            
        End With
        
        Dim i As Integer: For i = PIANOROLL_START_ROW To PIANOROLL_END_ROW
            With ThisWorkbook.Worksheets("ピアノロール")
                If .Cells(i, Counter).Borders(xlEdgeLeft).LineStyle = xlContinuous And .Cells(i, Counter).Interior.Color = NoteColor Then
                    Call Play_Sound(i - 1)
                ElseIf .Cells(i, Counter).Borders(xlEdgeRight).LineStyle = xlContinuous And .Cells(i, Counter).Interior.Color = NoteColor Then
                    Call Stop_Sound(i - 1)
                End If
            End With
        Next i
        
        'スクロール
        Windows(ThisWorkbook.Name).SmallScroll Down:=0, Up:=0, ToRight:=1, ToLeft:=0
        
        DoEvents
        
        t2 = Timer
        
        Sleep_Time = CLng((OneCellSeconds - (t2 - t1)) * 920)
        
        'Sleepで固まるのを防止するため
        If Sleep_Time < 0 Then Exit Do
        
        Sleep (Sleep_Time)
        
        Counter = Counter + 1
    Loop
   
End Sub

Public Sub End_Play()
    Dim ret As Long
    
    ret = midiOutClose(Handle)
    
End Sub



