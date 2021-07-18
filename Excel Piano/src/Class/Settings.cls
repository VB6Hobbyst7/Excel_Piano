VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'各種設定やプロパティを読み込み設定する。

Public Function Get_NoteColor() As Long
    '----------------------------------------------------------------
    '@Name
    '   Get_NoteColor
    '
    '@Param
    '   なし。
    '
    '@Return
    '   直接取得したRGB値を返す。
    '
    '@Description
    '   ノーツの色をRGB値を長整数型で返す。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    Get_NoteColor = CLng(ThisWorkbook.Worksheets("設定").Cells(3, 3).Value)
    
    
End Function

Public Function Get_EndColor() As Long
    '----------------------------------------------------------------
    '@Name
    '   Get_EndColor
    '
    '@Param
    '   なし。
    '
    '@Return
    '   直接取得したRGB値を返す。
    '
    '@Description
    '   エンドマーカーの色をRGB値を長整数型で返す。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    Get_EndColor = CLng(ThisWorkbook.Worksheets("設定").Cells(4, 3).Value)
    
    
End Function


Public Function Get_ScoreLength() As String
    '----------------------------------------------------------------
    '@Name
    '   Get_ScoreLength
    '
    '@Param
    '   なし。
    '
    '@Return
    '   直接取得した音符の最小単位を返す。
    '
    '@Description
    '   音符の最小単位を取得する。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    Get_ScoreLength = CStr(ThisWorkbook.Worksheets("設定").Cells(7, 3).Value)
    
End Function

Public Function Get_Tempo() As Integer
    '----------------------------------------------------------------
    '@Name
    '   Get_Tempo
    '
    '@Param
    '   なし。
    '
    '@Return
    '   直接取得した速度を返す。
    '
    '@Description
    '   曲の速さ(=テンポ)を取得する
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    Get_Tempo = CInt(ThisWorkbook.Worksheets("設定").Cells(8, 3).Value)
    
End Function

Public Function Get_Rhythm() As String
    '----------------------------------------------------------------
    '@Name
    '   Get_Rhythm
    '
    '@Param
    '   なし。
    '
    '@Return
    '   直接取得した拍子を返す。
    '
    '@Description
    '   拍子を取得する
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    Get_Rhythm = CStr(ThisWorkbook.Worksheets("設定").Cells(9, 3).Value)
    
End Function

Public Function Get_Title() As String
    '----------------------------------------------------------------
    '@Name
    '   Get_Title
    '
    '@Param
    '   なし。
    '
    '@Return
    '   直接曲名を返す。
    '
    '@Description
    '   曲名を読み込む。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    Get_Title = CStr(ThisWorkbook.Worksheets("設定").Cells(12, 3).Value)
    
End Function

Public Function Get_SavePath() As String
    '----------------------------------------------------------------
    '@Name
    '   Get_SavePath
    '
    '@Param
    '   なし。
    '
    '@Return
    '   直接保存先を返す。
    '
    '@Description
    '   保存先パスを読み込む。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    Get_SavePath = CStr(ThisWorkbook.Worksheets("設定").Cells(13, 3).Value)
    
End Function

Public Function Get_UseInstrument()
    '----------------------------------------------------------------
    '@Name
    '   Get_UseInstrument
    '
    '@Param
    '   なし。
    '
    '@Return
    '   楽器名が返る。
    '
    '@Description
    '   出力に使う楽器名を取得。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    Get_UseInstrument = CStr(ThisWorkbook.Worksheets("設定").Cells(18, 3).Value)
    
End Function

Public Sub Save_NoteColor(ByVal lngColor As Long)
    '----------------------------------------------------------------
    '@Name
    '   Save_NoteColor
    '
    '@Param
    '   lngColor    : 保存する色。
    '
    '@Return
    '   なし。
    '
    '@Description
    '   ノーツの色をRGB値を長整数型で設定シートに保存。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    ThisWorkbook.Worksheets("設定").Cells(3, 3).Value = lngColor
    
End Sub

Public Sub Save_EndColor(ByVal lngColor As Long)
     '----------------------------------------------------------------
    '@Name
    '   Save_EndColor
    '
    '@Param
    '   lngColor    : 保存する色。
    '
    '@Return
    '   なし。
    '
    '@Description
    '   エンドマーカーの色をRGB値を長整数型で設定シートに保存。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    ThisWorkbook.Worksheets("設定").Cells(4, 3).Value = lngColor
    
End Sub

Public Sub Save_ScoreLength(ByVal ScoreLength As String)
    '----------------------------------------------------------------
    '@Name
    '   Save_ScoreLength
    '
    '@Param
    '   ScoreLength     : 音符の最小単位。
    '
    '@Return
    '   なし。
    '
    '@Description
    '   音符の最小単位を設定シートに保存。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    ThisWorkbook.Worksheets("設定").Cells(7, 3).Value = ScoreLength
    
End Sub

Public Sub Save_Tempo(ByVal Tempo As Integer)
    '----------------------------------------------------------------
    '@Name
    '   Save_Tempo
    '
    '@Param
    '   Tempo   : テンポ。
    '
    '@Return
    '   なし。
    '
    '@Description
    '   テンポを設定シートに保存。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    ThisWorkbook.Worksheets("設定").Cells(8, 3).Value = Tempo
    
End Sub

Public Sub Save_Rhythm(ByVal Rhythm As String)
    '----------------------------------------------------------------
    '@Name
    '   Save_Rhythm
    '
    '@Param
    '   Rhythm      : 拍子。
    '
    '@Return
    '   なし。
    '
    '@Description
    '   拍子を設定シートに保存。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    ThisWorkbook.Worksheets("設定").Cells(9, 3).Value = Rhythm
    
End Sub

Public Sub Save_Title(ByVal Title As String)
    '----------------------------------------------------------------
    '@Name
    '   Save_Title
    '
    '@Param
    '   Title       : タイトル。
    '
    '@Return
    '   なし。
    '
    '@Description
    '   タイトルを設定シートに保存。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    ThisWorkbook.Worksheets("設定").Cells(12, 3).Value = Title
    
End Sub

Public Sub Save_SavePath(ByVal SavePath As String)
    '----------------------------------------------------------------
    '@Name
    '   Save_SavePath
    '
    '@Param
    '   SavePath    : 保存する色。
    '
    '@Return
    '   なし。
    '
    '@Description
    '   保存先を設定シートに保存。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    ThisWorkbook.Worksheets("設定").Cells(13, 3).Value = SavePath
    
End Sub


Public Function Save_UseInstrument(ByVal Ins As String)
    '----------------------------------------------------------------
    '@Name
    '   Save_UseInstrument
    '
    '@Param
    '   Ins         : 楽器名。
    '
    '@Return
    '   なし。
    '
    '@Description
    '   出力に使う楽器名を設定シートに保存。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
End Function


Public Function Get_OneCellSeconds(ByVal BPM As Integer, ScoreType As String) As Double
    '----------------------------------------------------------------
    '@Name
    '   Get_OneCellSeconds
    '
    '@Param
    '   BPM             : 演奏速度。
    '   ScoreType       : 音符の最小単位。
    '
    '@Return
    '   直接値を代入    : 1セルが占める秒数返す
    '
    '@Description
    '   音を重ね合わせて和音にするための関数。
    '
    '@Note
    '   120BPM→1分間に4分音符が120個
    '   16分音符が最小とすると■■■■で4分音符(■は16分音符)
    '   60秒で120個→1秒で2個。つまり1セルあたり0.5秒である。
    '   同様にしてBPM = K の時、1セルあたりの秒数は(60 / K) / (音符固有の値)となる
    '
    '----------------------------------------------------------------
    
    '音符の最小単位から音符固有の値を決める
    Dim div As Double
    If ScoreType = "全" Then
        div = 1 / 4
    ElseIf ScoreType = "2分" Then
        div = 1 / 2
    ElseIf ScoreType = "4分" Then
        div = 1
    ElseIf ScoreType = "8分" Then
        div = 2
    ElseIf ScoreType = "16分" Then
        div = 4
    ElseIf ScoreType = "32分" Then
        div = 8
    ElseIf ScoreType = "64分" Then
        div = 16
    Else
        div = 4
    End If
    
    Get_OneCellSeconds = ((60 / BPM) / div)
    
End Function

Public Function Get_LastColumn() As Long
    '----------------------------------------------------------------
    '@Name
    '   Get_LastColumn
    '
    '@Param
    '   なし。
    '
    '@Return
    '   LastPos     : 最終列の列番号。
    '
    '@Description
    '   曲の終わりの列を取得する。見つからない場合は-1を返す。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    Dim LastPos As Long
    Dim Is_Find As Boolean: Is_Find = False
    
    Dim i As Long
    For i = PIANOROLL_START_COLUMN To Columns.Count
        If ThisWorkbook.Worksheets("ピアノロール").Cells(PIANOROLL_START_ROW, i).Interior.Color = Get_EndColor() Then
            Is_Find = True
            LastPos = i
            Exit For
        End If
    Next i
    
    If Is_Find = False Then LastPos = -1
    
    
    Get_LastColumn = LastPos
    
    
End Function

Public Function Is_In_Scope(ByVal TargetRange As Range) As Boolean
    '----------------------------------------------------------------
    '@Name
    '   Is_In_Scope
    '
    '@Param
    '   TargetRange     : 選択範囲。
    '
    '@Return
    '   Result          : 0で有効。それ以外は無効。
    '
    '@Description
    '   選択範囲がノーツを入力するのに有効かどうかをBoolean型で返す。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------

    Dim StartRow As Long: StartRow = TargetRange.Rows(1).Row
    Dim EndRow As Long: EndRow = TargetRange.Rows(TargetRange.Rows.Count).Row
    Dim StartColumn As Long: StartColumn = TargetRange.Columns(1).Column
    Dim EndColumn As Long: EndColumn = TargetRange.Columns(TargetRange.Columns.Count).Column
    Dim Result As Long: Result = 0
    
    
    If StartRow = NOTE_LENGTH_ROW Then Result = Result + 1
    If StartColumn = KEYBOARD_COLUMN Then Result = Result + 1
    If EndRow = KEYBOARD_COLUMN Then Result = Result + 1
    If EndRow > PIANOROLL_END_ROW Then Result = Result + 1
    If EndColumn = 1 Then Result = Result + 1
    
    Is_In_Scope = Not CBool(Result)
    
End Function

Public Function Get_ScoreLength_Assist(ByVal ScoreType As String) As Long
    '----------------------------------------------------------------
    '@Name
    '   Get_ScoreLength_Assist
    '
    '@Param
    '   ScoreType       : 入力アシストに使う音符の種類。
    '
    '@Return
    '   CellLength      : 入力に使うセルの長さ。
    '
    '@Description
    '   入力アシストに使う音符の種類から打ち込むセルの長さを計算。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    Dim CellLength As Long
    Dim ScoreLength As String: ScoreLength = Get_ScoreLength()
    Dim WholeCellCount As Integer  '全音符が使うセルの数
    
    '一旦全分音符が使うセルの数をセット
    Select Case ScoreLength
        Case "全"
            WholeCellCount = 1
        Case "2分"
            WholeCellCount = 2
        Case "4分"
            WholeCellCount = 4
        Case "8分"
            WholeCellCount = 8
        Case "16分"
            WholeCellCount = 16
        Case "32分"
            WholeCellCount = 32
        Case "64分"
            WholeCellCount = 64
    End Select
    
    Select Case ScoreType
        Case "全音符"
            CellLength = WholeCellCount * 1
        Case "2分音符"
            CellLength = WholeCellCount / 2
        Case "4分音符"
            CellLength = WholeCellCount / 4
        Case "8分音符"
            CellLength = WholeCellCount / 8
        Case "16分音符"
            CellLength = WholeCellCount / 16
        Case "付点全音符"
            CellLength = WholeCellCount / 1 + WholeCellCount / 2
        Case "付点2分音符"
            CellLength = WholeCellCount / 2 + WholeCellCount / 4
        Case "付点4分音符"
            CellLength = WholeCellCount / 4 + WholeCellCount / 8
        Case "付点8分音符"
            CellLength = WholeCellCount / 8 + WholeCellCount / 16
        Case "付点16分音符"
            CellLength = WholeCellCount / 16 + WholeCellCount / 32
    End Select
    
    Get_ScoreLength_Assist = CellLength
    
End Function

Public Sub Save_Settings_All()
    '----------------------------------------------------------------
    '@Name
    '   Save_Settings
    '
    '@Param
    '   なし。
    '
    '@Return
    '   なし。
    '
    '@Description
    '   設定情報を保存する。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    '以下、設定シートの項目上から順に保存
    With ThisWorkbook.Worksheets("設定")
        .Cells(3, 3).Value = MainForm.NoteColor_Label.BackColor 'ノートカラー
        .Cells(4, 3).Value = MainForm.EndColor_Label.BackColor  'エンドマークカラー
        .Cells(7, 3).Value = MainForm.ScoreLength_ComboBox.Text '音符の最小単位
        .Cells(8, 3).Value = MainForm.Tempo_TextBox.Text        '速度(テンポ)
        .Cells(9, 3).Value = MainForm.Rhythm_ComboBox.Text      '拍子
        .Cells(12, 3).Value = MainForm.Title_TextBox.Text       'タイトル(曲名)
        .Cells(18, 3).Value = MainForm.Instrument_ComboBox.Text '使用楽器
    End With
    
    
End Sub

Public Sub Set_ComboBox_InstrumentList(ByRef TargetComboBox As ComboBox)
    '----------------------------------------------------------------
    '@Name
    '   Set_ComboBox_InstrumentList
    '
    '@Param
    '   TargetComboBox      : 楽器リストを設定するコンボボックス。
    '
    '@Return
    '   なし。
    '
    '@Description
    '   SoundSourceフォルダにある楽器名を取得しコンボボックスにセットする。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    Dim strPattern As String, strInstrument As String
    strPattern = Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\", InStrRev(ThisWorkbook.Path, "\") - 1) - 1) + "\SoundSource\"
     
    strInstrument = Dir(strPattern, vbDirectory)

    Do While Len(strInstrument) > 0
        If GetAttr(strPattern & strInstrument) And vbDirectory Then
            If strInstrument <> "." And strInstrument <> ".." Then
                TargetComboBox.AddItem (strInstrument)
            End If
        End If
        strInstrument = Dir()
    Loop
    
    
End Sub


Public Sub Set_ComboBox_ScoreLength(ByRef TargetComboBox As ComboBox)
    '----------------------------------------------------------------
    '@Name
    '   Set_ComboBox_ScoreLength
    '
    '@Param
    '   TargetComboBox : 項目を設定するコンボボックス。
    '
    '@Return
    '   なし。直接コンボボックスに項目が設定される。
    '
    '@Description
    '   音符の最小単位(16分音符とか)を指定のコンボボックスに設定する。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    Dim Origin As Range: Set Origin = ThisWorkbook.Worksheets("設定").Cells(16, 2)
    
    Dim i As Long
    For i = 1 To 7
        TargetComboBox.AddItem (CStr(Origin.Offset(0, i).Value))
    Next i
    
    'デフォルトを表示
    TargetComboBox.Text = CStr(Origin.Offset(0, 5).Value)
    
End Sub

Public Sub Set_ComboBox_Rhythm(ByRef TargetComboBox As ComboBox)
    '----------------------------------------------------------------
    '@Name
    '   Set_ComboBox_ScoreLength
    '
    '@Param
    '   TargetComboBox : 項目を設定するコンボボックス。
    '
    '@Return
    '   なし。直接コンボボックスに項目が設定される。
    '
    '@Description
    '   拍子(4/4拍子とか)を指定のコンボボックスに設定する。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    Dim Origin As Range: Set Origin = ThisWorkbook.Worksheets("設定").Cells(17, 2)
    
    Dim i As Long
    For i = 1 To 6
        TargetComboBox.AddItem (CStr(Origin.Offset(0, i).Value))
    Next i
    
    'デフォルトを表示
    TargetComboBox.Text = CStr(Origin.Offset(0, 2).Value)
End Sub

