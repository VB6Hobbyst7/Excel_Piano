Attribute VB_Name = "Variable"
Option Explicit

'グローバル変数

Public Const PIANOROLL_START_ROW As Integer = 2         '鍵盤の始まりの行
Public Const PIANOROLL_END_ROW As Integer = 89          '鍵盤の終わりの行
Public Const PIANOROLL_START_COLUMN As Integer = 2      '鍵盤の始まりの列
Public PIANOROLL_END_COLUMN As Integer                  '鍵盤の終わりの列
Public Const NOTE_LENGTH_ROW As Integer = 1             '連符用の列
Public Const SUSTAIN_ROW As Integer = 90                'サステイン用の列
Public Const KEYBOARD_COLUMN As Integer = 1             '鍵盤が描かれている列

Public Mode As String           '入力モード(選択、追加、削除、分割、結合)

Public AssistMode As String     'アシストモード

Public Threshold As Integer     '入力アシストの無効化を判断する閾値

Public NoteOrEnd As String      'ノートかエンドマーカーか入力する際に使う

Public Is_Playing As Boolean    'リアルタイム演奏中かどうか(演奏中でTrue)

Private CSettings As New Settings

'全部のノーツの80を入れる
Sub Set_Volume()
    Dim i As Long, j As Long
    With ThisWorkbook.Worksheets("ピアノロール")
        For i = PIANOROLL_START_ROW To PIANOROLL_END_ROW
            For j = PIANOROLL_START_COLUMN To CSettings.Get_LastColumn
                If .Cells(i, j).Interior.Color = CSettings.Get_NoteColor() Then
                    If .Cells(i, j).Borders(xlEdgeLeft).LineStyle = xlContinuous Then
                        .Cells(i, j).Value = 80
                    End If
                End If
            Next j
        Next i
    End With
End Sub

Sub AllClear()
    Range(Cells(PIANOROLL_START_ROW, PIANOROLL_START_COLUMN), Cells(PIANOROLL_END_ROW, Columns.Count)).Clear
End Sub

'いい感じにサステインを設定できるかも?
Sub SetSustain()
    
    Const Interval As Integer = 8   '4/4 最小16分音符の時に1小節ごとにサステインを入れる
    
    Dim i As Long
    Dim OneCellSeconds As Double, BasicOneCellSeconds As Double
    BasicOneCellSeconds = CSettings.Get_OneCellSeconds(CSettings.Get_Tempo, CSettings.Get_ScoreLength)
    Dim StartColumn As Long: StartColumn = 6
    
    Dim tmp As Double
    For i = StartColumn To CSettings.Get_LastColumn
    
        '連符があるか調べ、1セル当たりの秒数を変える
        If Cells(NOTE_LENGTH_ROW, i).Value <> "" Then
            OneCellSeconds = 2 * BasicOneCellSeconds / CInt(Cells(NOTE_LENGTH_ROW, i).Value)
        Else
            OneCellSeconds = BasicOneCellSeconds
        End If
        
        
        tmp = tmp + OneCellSeconds
        
        If Abs(tmp - BasicOneCellSeconds * Interval) < 0.01 Then
            Cells(SUSTAIN_ROW, i).Value = "E"
            Cells(SUSTAIN_ROW, i + 1).Value = "S"
            tmp = 0
        End If
        
        
    Next i
End Sub


