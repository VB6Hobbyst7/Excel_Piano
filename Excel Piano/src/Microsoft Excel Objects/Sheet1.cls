VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private CSettings As New Settings

'ダブルクリックでMainFormを開く
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    
    MainForm.Show vbModeless
    
    '入力中にしない
    Cancel = True
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    'モードが「Add」以外は除く
    If Mode <> "Add" Then Exit Sub
    
    Dim StartRow As Long: StartRow = Target.Rows(1).Row
    Dim EndRow As Long: EndRow = Target.Rows(Target.Rows.Count).Row
    Dim StartColumn As Long: StartColumn = Target.Columns(1).Column
    Dim EndColumn As Long: EndColumn = Target.Columns(Target.Columns.Count).Column
    
    '列範囲選択ははじく
    If EndRow - StartRow > 1 And EndColumn - StartColumn > 1 Then Exit Sub
    
    '選択する長さを取得
    Dim Select_Length As Long: Select_Length = CSettings.Get_ScoreLength_Assist(AssistMode)
    ThisWorkbook.Worksheets("ピアノロール").Range(Cells(StartRow, StartColumn), Cells(StartRow, StartColumn + Select_Length - 1)).Select
    
    Cancel = True
    
End Sub

'ベロシティーの値が変更されたときに発生
Private Sub Worksheet_Change(ByVal Target As Range)
    
    'イベントを捕捉しない
    Application.EnableEvents = False
    
    Dim StartRow As Long, EndRow As Long
    Dim StartColumn As Long, EndColumn As Long
    Dim BeforeChangeCell As Range
    
    StartRow = Target(1).Row
    EndRow = Target(Target.Count).Row
    StartColumn = Target(1).Column
    EndColumn = Target(Target.Count).Column
  
    
    If EndColumn - StartColumn <> 0 Or EndRow - StartRow <> 0 Then Application.EnableEvents = True: Exit Sub
    If StartColumn = KEYBOARD_COLUMN Or StartRow = NOTE_LENGTH_ROW Or EndRow > PIANOROLL_END_ROW Then Application.EnableEvents = True: Exit Sub
    
    
    If Target.Value = "" Or IsNumeric(Target.Value) = False Then Application.EnableEvents = True: Exit Sub
    Dim Volume As Integer: Volume = CInt(Target.Value)
    If Volume < 0 Or Volume > 100 Then Volume = 80
    
    With ThisWorkbook.Worksheets("ピアノロール")
        'ノーツの部分に入力されていたときのみ実行
        If .Cells(StartRow, StartColumn).Interior.Color = CSettings.Get_NoteColor() Then
            'ノーツの開始点じゃなかったら最初に戻って入力
            If .Cells(StartRow, StartColumn).Borders(xlEdgeLeft).LineStyle <> xlContinuous Then
                Dim BeginCellColumn As Long: BeginCellColumn = StartColumn
                Do Until (.Cells(StartRow, BeginCellColumn).Borders(xlEdgeLeft).LineStyle = xlContinuous Or BeginCellColumn <= PIANOROLL_START_COLUMN)
                    .Cells(StartRow, BeginCellColumn).ClearContents
                    BeginCellColumn = BeginCellColumn - 1
                Loop
                .Cells(StartRow, BeginCellColumn).Value = Volume
            End If
        Else
            '入力モードとして入れられた数値から音符の種類を判断
            Application.EnableEvents = True
            Dim Mode_Backup As String: Mode_Backup = Mode
            Mode = "Add"
            '選択する長さを取得
            Dim Add_Score As String: Add_Score = "4分音符"
            Select Case Target.Value
                Case 0
                    Add_Score = "全音符"
                Case 0.5
                    Add_Score = "付点全音符"
                Case 2
                    Add_Score = "2分音符"
                Case 2.5
                    Add_Score = "付点2分音符"
                Case 4
                    Add_Score = "4分音符"
                Case 4.5
                    Add_Score = "付点4分音符"
                Case 8
                    Add_Score = "8分音符"
                Case 8.5
                    Add_Score = "付点8分音符"
                Case 16
                    Add_Score = "16分音符"
                Case 16.5
                    Add_Score = "付点16分音符"
            End Select
            Dim Select_Length As Long: Select_Length = CSettings.Get_ScoreLength_Assist(Add_Score)
            ThisWorkbook.Worksheets("ピアノロール").Range(Cells(StartRow, StartColumn), Cells(StartRow, StartColumn + Select_Length - 1)).Select
            Mode = Mode_Backup
        End If
        
    End With
    
    
    Application.EnableEvents = True
    
End Sub

'選択範囲が変更されたときに発生(入力用)
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If CSettings.Is_In_Scope(Target) = False Then Exit Sub
    
    
    Dim StartRow As Long, EndRow As Long
    Dim StartColumn As Long, EndColumn As Long
    StartRow = Selection(1).Row
    EndRow = Selection(Selection.Count).Row
    StartColumn = Selection(1).Column
    EndColumn = Selection(Selection.Count).Column
    
    
    'エンドマーカーの入力の時
    If Mode = "Add" And NoteOrEnd = "End" Then
        '範囲選択は除く
        If EndColumn - StartColumn <> 0 Or EndRow - StartRow <> 0 Then Exit Sub
        With ThisWorkbook.Worksheets("ピアノロール")
            PIANOROLL_END_COLUMN = CSettings.Get_LastColumn
            If PIANOROLL_END_COLUMN = -1 Then   'そもそも終わりが設定されていないとき
                .Cells(PIANOROLL_START_ROW, StartColumn).Interior.Color = CSettings.Get_EndColor
            Else    '終わりが存在している場合
                .Cells(PIANOROLL_START_ROW, PIANOROLL_END_COLUMN).Clear
                .Cells(PIANOROLL_START_ROW, StartColumn).Interior.Color = CSettings.Get_EndColor
            End If
        End With
        
        Exit Sub
    End If
    
    'ノーツの入力の時
    If Mode = "Select" Then

    ElseIf Mode = "Add" Then
        Dim DefaultVolume As Integer: DefaultVolume = 80
        With ThisWorkbook.Worksheets("ピアノロール")
            '一旦すべてクリア
            .Range(Cells(StartRow, StartColumn), Cells(EndRow, EndColumn)).ClearFormats
            
            Dim i As Integer
            
            For i = StartRow To EndRow
                Range(Cells(i, StartColumn), Cells(i, EndColumn)).Interior.Color = CSettings.Get_NoteColor()
                '罫線を引いて領域を明確にする
                .Range(Cells(i, StartColumn), Cells(i, EndColumn)).Borders(xlEdgeTop).LineStyle = xlContinuous   '上の線
                .Range(Cells(i, StartColumn), Cells(i, EndColumn)).Borders(xlEdgeTop).Weight = xlThin
                
                .Range(Cells(i, StartColumn), Cells(i, EndColumn)).Borders(xlEdgeBottom).LineStyle = xlContinuous    '下の線
                .Range(Cells(i, StartColumn), Cells(i, EndColumn)).Borders(xlEdgeBottom).Weight = xlThin
                
                .Cells(i, EndColumn).Borders(xlEdgeRight).LineStyle = xlContinuous   '右の線
                .Cells(i, EndColumn).Borders(xlEdgeRight).Weight = xlThin
                
                .Cells(i, StartColumn).Borders(xlEdgeLeft).LineStyle = xlContinuous  '左の線
                .Cells(i, StartColumn).Borders(xlEdgeLeft).Weight = xlThin
                
                .Cells(i, StartColumn).Value = DefaultVolume
                
            Next i
            
        End With
        
    ElseIf Mode = "Delete" Then
        With ThisWorkbook.Worksheets("ピアノロール")
            '指定の範囲をクリアする
            .Range(Cells(StartRow, StartColumn), Cells(EndRow, EndColumn)).Clear
            
            Dim j As Integer
            
            For j = StartRow To EndRow
                If .Cells(j, StartColumn - 1).Interior.Color = CSettings.Get_NoteColor() And .Cells(j, StartColumn - 1).Borders(xlEdgeRight).LineStyle = xlLineStyleNone Then
                    .Cells(j, StartColumn - 1).Borders(xlEdgeRight).LineStyle = xlContinuous
                    .Cells(j, StartColumn - 1).Borders(xlEdgeRight).Weight = xlThin
                End If
                
                If .Cells(j, EndColumn + 1).Interior.Color = CSettings.Get_NoteColor() And Cells(j, EndColumn + 1).Borders(xlEdgeLeft).LineStyle = xlLineStyleNone Then
                    .Cells(j, EndColumn + 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
                    .Cells(j, EndColumn + 1).Borders(xlEdgeLeft).Weight = xlThin
                End If
            Next j
        End With
        
    ElseIf Mode = "Divide" Then
        With ThisWorkbook.Worksheets("ピアノロール")
            '範囲選択は除く
            If EndColumn - StartColumn <> 0 Or EndRow - StartRow <> 0 Then Exit Sub
            
             '選択したセルの右側にノートがあれば線を追加する
            If .Cells(StartRow, StartColumn).Interior.Color = CSettings.Get_NoteColor() And .Cells(StartRow, StartColumn + 1).Interior.Color = CSettings.Get_NoteColor() Then
                .Cells(StartRow, StartColumn).Borders(xlEdgeRight).LineStyle = xlContinuous
                .Cells(StartRow, StartColumn).Borders(xlEdgeRight).Weight = xlThin
                
                .Cells(StartRow, StartColumn + 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
                .Cells(StartRow, StartColumn + 1).Borders(xlEdgeLeft).Weight = xlThin
                
                .Cells(StartRow, StartColumn + 1).Value = 80
            End If
        End With
        
    ElseIf Mode = "Join" Then
        With ThisWorkbook.Worksheets("ピアノロール")
            '範囲選択は除く
            If EndColumn - StartColumn <> 0 Or EndRow - StartRow <> 0 Then Exit Sub
            
             '選択したセルの罫線、または選択したセルの左側の罫線が「xlEdgeRight」になっていれば結合
            If .Cells(StartRow, StartColumn).Borders(xlEdgeRight).LineStyle = xlContinuous And .Cells(StartRow, StartColumn + 1).Interior.Color = CSettings.Get_NoteColor() Then
                .Cells(StartRow, StartColumn).Borders(xlEdgeRight).LineStyle = xlLineStyleNone
                .Cells(StartRow, StartColumn + 1).Value = ""
                
            ElseIf .Cells(StartRow, StartColumn - 1).Borders(xlEdgeRight).LineStyle = xlContinuous And .Cells(StartRow, StartColumn + 1).Interior.Color = CSettings.Get_NoteColor() Then
                .Cells(StartRow, StartColumn - 1).Borders(xlEdgeRight).LineStyle = xlLineStyleNone
                .Cells(StartRow, StartColumn + 1).Value = ""
                
            End If
        End With
        
    End If



End Sub

