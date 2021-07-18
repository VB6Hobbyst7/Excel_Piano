VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScoreData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'楽譜ファイルをインポートしたりエクスポートしたりするクラス

Private CSettings As New Settings

'ピアノロールに楽譜データをインポートする
Public Sub Import_ScoreData_EMD(ByVal Path As String)
    '----------------------------------------------------------------
    '@Name
    '   Import_ScoreData_EMD
    '
    '@Param
    '   Path        : インポート先の絶対パス。
    '
    '@Return
    '   なし。
    '
    '@Description
    '   EMD形式の楽譜データを読み込んで表示する。
    '
    '@Note
    '   配列に入っている順番は
    '   開始行, 連符, 開始列, ベロシティ, 終了列
    '
    '----------------------------------------------------------------
    
    With ThisWorkbook.Worksheets("ピアノロール")
        .Range(.Cells(NOTE_LENGTH_ROW, PIANOROLL_START_COLUMN), .Cells(NOTE_LENGTH_ROW, Columns.Count)).ClearContents '連符を削除
        .Range(.Cells(SUSTAIN_ROW, PIANOROLL_START_COLUMN), .Cells(SUSTAIN_ROW, Columns.Count)).ClearContents
        .Range(.Cells(PIANOROLL_START_ROW, PIANOROLL_START_COLUMN), .Cells(PIANOROLL_END_ROW, Columns.Count)).Clear   'ピアノロール内のデータすべてを消す
    
    
        Dim ScoreData() As String
        
        'ノーツとサステインの区切れの行
        Dim SeparationRow As Long
        
        'ファイルを開いて情報の取得
        Dim n As Long: n = 0
        Open Path For Input As #1
        
            Do While Not EOF(1)
                '読み込み
                ReDim Preserve ScoreData(n)
                '一行ずつ読み込む
                Line Input #1, ScoreData(n)
                
                'ノーツとサステインデータの区切れか調べる
                If ScoreData(n) = String(30, "-") Then SeparationRow = n
                
                'インクリメント
                n = n + 1
            Loop
            
        Close #1
        
        
        '取得した情報をもとにデータを配置
        Call CSettings.Save_Title(ScoreData(0))  '曲名
        Call CSettings.Save_Tempo(CInt(ScoreData(1)))   'テンポ
        Call CSettings.Save_ScoreLength(ScoreData(2))   '音符の最小単位
        Call CSettings.Save_Rhythm(ScoreData(3))        '拍子
        .Cells(PIANOROLL_START_ROW, CLng(ScoreData(4))).Interior.Color = CSettings.Get_EndColor   '終わりを設定

        For n = 6 To SeparationRow - 1 '5番目の要素は区切り
            '読み込んだデータを「,」で区切る
            Dim Note_Data() As String: Note_Data = Split(ScoreData(n), ",") '配列の要素数は0,1,2,3,4の5つ
            
            '色を付ける
            .Range(.Cells(CLng(Note_Data(0)), CLng(Note_Data(2))), .Cells(CLng(Note_Data(0)), CLng(Note_Data(4)))).Interior.Color = CSettings.Get_NoteColor
            
            '上線と下線を引く
            .Range(.Cells(CLng(Note_Data(0)), CLng(Note_Data(2))), .Cells(CLng(Note_Data(0)), CLng(Note_Data(4)))).Borders(xlEdgeTop).LineStyle = xlContinuous
            .Range(.Cells(CLng(Note_Data(0)), CLng(Note_Data(2))), .Cells(CLng(Note_Data(0)), CLng(Note_Data(4)))).Borders(xlEdgeTop).Weight = xlThin
            .Range(.Cells(CLng(Note_Data(0)), CLng(Note_Data(2))), .Cells(CLng(Note_Data(0)), CLng(Note_Data(4)))).Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Range(.Cells(CLng(Note_Data(0)), CLng(Note_Data(2))), .Cells(CLng(Note_Data(0)), CLng(Note_Data(4)))).Borders(xlEdgeBottom).Weight = xlThin
            
            '始まりに左線
            .Cells(CLng(Note_Data(0)), CLng(Note_Data(2))).Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Cells(CLng(Note_Data(0)), CLng(Note_Data(2))).Borders(xlEdgeLeft).Weight = xlThin
            
            '終わりに右線
            .Cells(CLng(Note_Data(0)), CLng(Note_Data(4))).Borders(xlEdgeRight).LineStyle = xlContinuous
            .Cells(CLng(Note_Data(0)), CLng(Note_Data(4))).Borders(xlEdgeRight).Weight = xlThin
            
            '速度に変更がある場合、反映
            If Note_Data(1) <> "" Then
                .Range(.Cells(NOTE_LENGTH_ROW, CLng(Note_Data(2))), .Cells(NOTE_LENGTH_ROW, CLng(Note_Data(4)))).Value = Note_Data(1)
            End If
            
            '音の大きさを設定
            .Cells(CLng(Note_Data(0)), CLng(Note_Data(2))).Value = Note_Data(3)
            
        Next n
        
        
        For n = SeparationRow + 1 To UBound(ScoreData)
            Dim Sustain_Data() As String: Sustain_Data = Split(ScoreData(n), ",")
            
            .Cells(SUSTAIN_ROW, CLng(Sustain_Data(0))).Value = Sustain_Data(1)
        Next n
    
    End With
    
End Sub

'ピアノロールから楽譜データを書き出す
Public Sub Export_ScoreData_EMD(ByVal Path As String)
    '----------------------------------------------------------------
    '@Name
    '   Export_ScoreData_EMD
    '
    '@Param
    '   Path        : インポート先の絶対パス。
    '
    '@Return
    '   なし。
    '
    '@Description
    '   EMD形式の楽譜データを出力する。
    '
    '@Note
    '   配列に入っている順番は
    '   開始行, 連符, 開始列, ベロシティ, 終了列
    '
    '----------------------------------------------------------------
    
    
    '終わりがない場合は終了
    If PIANOROLL_END_COLUMN = -1 Then MsgBox ("曲の終わりを設定してください。"): Exit Sub
    
    
    Dim Title As String     '曲名
    Dim Tempo As Integer    '速度
    Dim ScoreKind As String '音符の最小単位
    Dim Rhythm As String    '拍子
    
    '基本情報の取得
    Title = CSettings.Get_Title
    Tempo = CSettings.Get_Tempo
    ScoreKind = CSettings.Get_ScoreLength
    Rhythm = CSettings.Get_Rhythm
    
    '楽譜データを配列化する
    Dim ScoreData() As String
    Dim FlgStart As Boolean '始まりがあったかどうか
    Dim Counter As Long     '音がある場所が何箇所あったか(最終出力はヘッダを除いて5からスタート)
    
    '変数の初期化
    FlgStart = False
    Counter = 6
    
    ReDim ScoreData(5) As String
    ScoreData(0) = Title    '曲名
    ScoreData(1) = CStr(Tempo)  '速度
    ScoreData(2) = ScoreKind    '最小の音符の種類
    ScoreData(3) = CStr(Rhythm) '拍子
    ScoreData(4) = CStr(PIANOROLL_END_COLUMN) '最後を示すセルがある列
    ScoreData(5) = String(30, "-")
    
    Dim i As Integer, j As Integer
    PIANOROLL_END_COLUMN = ScoreData(4)
    
    With ThisWorkbook.Worksheets("ピアノロール")
        
        'ノーツデータの入力
        For i = PIANOROLL_START_ROW To PIANOROLL_END_ROW
            
            For j = PIANOROLL_START_COLUMN To PIANOROLL_END_COLUMN
                'ノーツが始まるとき
                If .Cells(i, j).Interior.Color = CSettings.Get_NoteColor And .Cells(i, j).Borders(xlEdgeLeft).LineStyle = xlContinuous Then
                    'スタートチェッカーをTrueに
                    FlgStart = True
                    
                    '配列要素数の変更
                    ReDim Preserve ScoreData(Counter) As String
                    
                    'データの代入(開始行, 連符, 開始列, ベロシティ)
                    ScoreData(Counter) = CStr(CStr(i) + "," + CStr(.Cells(NOTE_LENGTH_ROW, j).Value) + "," + CStr(j) + "," + CStr(.Cells(i, j).Value) + ",")
                    
                End If
                
                'ノートが終わる場合
                If .Cells(i, j).Interior.Color = CSettings.Get_NoteColor And .Cells(i, j).Borders(xlEdgeRight).LineStyle = xlContinuous And FlgStart = True Then
                    'スタートチェッカーをFalseに
                    FlgStart = False
                    
                    '終わりの列を代入
                    ScoreData(Counter) = ScoreData(Counter) + CStr(j)
                    
                    Counter = Counter + 1
                End If
                
            Next j
            
        Next i
        
        'ノーツの終わりの区切り
        ReDim Preserve ScoreData(Counter) As String
        ScoreData(Counter) = String(30, "-")
        
        
        'サステイン情報を入力
        Counter = Counter + 1
        Dim c As Long: For c = PIANOROLL_START_COLUMN To PIANOROLL_END_COLUMN
            If UCase(.Cells(SUSTAIN_ROW, c).Value) = "S" Or UCase(.Cells(SUSTAIN_ROW, c).Value) = "E" Then
                ReDim Preserve ScoreData(Counter) As String
                ScoreData(Counter) = CStr(CStr(c) + "," + CStr(.Cells(SUSTAIN_ROW, c).Value))
                Counter = Counter + 1
            End If
        Next c
    
    End With
    
    'ファイルに書き出し(拡張子は.emd(つまりExcel Music Dataの略))
    'ファイルを開く(Outputモード)
    Dim n As Long
    Open Path For Output As #1
        For n = 0 To UBound(ScoreData)
            Print #1, ScoreData(n)
        Next n
    Close #1
    
    
End Sub


