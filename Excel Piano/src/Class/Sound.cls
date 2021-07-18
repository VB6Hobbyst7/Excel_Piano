VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'音の波形を計算する

Private CSettings As New Settings
Private CWav As New Wav

Private Const PI As Double = 3.14159265358979

Private Preset(87) As PresetData
Private Type PresetData
    WaveForm() As Integer
End Type

Public Sub Read_InstrumentSound(ByVal Instrument As String)
    '----------------------------------------------------------------
    '@Name
    '   Read_InstrumentSound
    '
    '@Param
    '   freqRow         : 音の高さを表す周波数がある行。
    '   Sampling        : サンプリング周波数。
    '   Seconds         : 音を生成する時間の長さ(秒)。
    '   Is_SoundBreak   : 音が途切れるかどうか
    '
    '@Return
    '   計算されたピアノ音の波形がDouble型の配列で返される。
    '
    '@Description
    '   任意の音階のピアノ音を生成する関数。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    Dim SoundDir As String: SoundDir = Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\", InStrRev(ThisWorkbook.Path, "\") - 1) - 1) + "\SoundSource\"
    Dim i As Integer
    For i = 0 To 87
        Call CWav.Read_Wave_16bits_Mono(SoundDir + Instrument + "\" + Instrument + "(" + CStr(i + 1) + ").wav", Preset(i).WaveForm)
    Next i
End Sub

Private Function Generate_Piano_Sound(ByVal freqRow As Long, ByVal Sampling As Long, ByVal Seconds As Double, ByVal Release_Time As Double, ByVal Volume As Integer) As Double()
    '----------------------------------------------------------------
    '@Name
    '   Generate_Piano_Sound
    '
    '@Param
    '   freqRow         : 音の高さを表す周波数がある行。
    '   Sampling        : サンプリング周波数。
    '   Seconds         : 音を生成する時間の長さ(秒)。
    '   Release_Time    : リリースタイム
    '   Volue           : ボリューム
    '
    '@Return
    '   計算されたピアノ音の波形がDouble型の配列で返される。
    '
    '@Description
    '   任意の音階のピアノ音を生成する関数。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    
    '指定秒数だけ配列から要素を抜き取る
    Dim Result() As Double: ReDim Result(Sampling * Seconds) As Double
    Dim i As Long
    For i = 0 To UBound(Result)
        Result(i) = Preset(freqRow - 2).WaveForm(i) * (Volume / 100)
    Next i
    
    Dim j, k As Long
    'フェード処理をかける
    For j = 0 To (Sampling * 0.01) - 1
        Result(j) = Result(j) * (CDbl(j) / (Sampling * 0.01))
    Next j
    
    For k = 0 To (Sampling * Release_Time) - 1
        Result(UBound(Result) - k - 1) = Result(UBound(Result) - k - 1) * (CDbl(k) / (Sampling * Release_Time))
    Next k
    
    
    Result(UBound(Result)) = 0
    
    Generate_Piano_Sound = Result
    
    
End Function

Private Sub Add_Sound(ByRef WaveFormData1() As Double, ByRef WaveFormData2() As Double, ByVal Pos As Long)
    '----------------------------------------------------------------
    '@Name
    '   Add_Sound
    '
    '@Param
    '   WaveFormData1   : 重ねられる波形データ。
    '   WaveFormData2   : 重ねる波形データ。
    '   Pos             : 重ね合わせ開始点。配列の添え字の部分を指定。
    '
    '@Return
    '   なし。引数の配列に入っている波形データが更新される。
    '
    '@Description
    '   音を重ね合わせて和音にするための関数。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    Dim Length As Long: Length = UBound(WaveFormData2)
    Dim i As Long
    For i = 0 To Length
        WaveFormData1(Pos + i) = CDbl(WaveFormData1(Pos + i)) + CDbl(WaveFormData2(i)) * 0.00001
    Next i
    
End Sub

Private Function Get_Data_Count(ByVal Sampling As Long, ByVal BPM As Integer, ScoreType As String) As Long
    '----------------------------------------------------------------
    '@Name
    '   Get_Data_Count
    '
    '@Param
    '   BPM         : 曲の速さ。
    '   ScoreType   : 音符の最小単位。
    '
    '@Return
    '   波形データが入るデータ個数を返す。
    '
    '@Description
    '   波形データが入るデータ個数を返す。
    '
    '@Note
    '   なし。
    '
    '----------------------------------------------------------------
    
    Dim BasicOneCellSeconds As Double, OneCellSeconds As Double
    BasicOneCellSeconds = CSettings.Get_OneCellSeconds(BPM, ScoreType)
    
    
    Dim i As Long
    Dim Counter As Long: Counter = 0
    For i = PIANOROLL_START_COLUMN To PIANOROLL_END_COLUMN
        With ThisWorkbook.Worksheets("ピアノロール")
            '連符があるか調べ、1セル当たりの秒数を設定
            If .Cells(NOTE_LENGTH_ROW, i).Value <> "" Then
                OneCellSeconds = 2 * BasicOneCellSeconds / CInt(.Cells(NOTE_LENGTH_ROW, i).Value)
            Else
                OneCellSeconds = BasicOneCellSeconds
            End If
             
            Counter = Counter + (OneCellSeconds * Sampling)
            
        End With
    Next i
    
    Get_Data_Count = Counter
    
End Function

Public Sub Export_Mixdown(ByVal FilePath As String, ByVal Sampling As Long, ByVal BPM As Integer, ScoreType As String, ByVal Instrument As String)
    '----------------------------------------------------------------
    '@Name
    '   Export_Mixdown
    '
    '@Param
    '   FilePath    : 書き出す保存先のファイルパス。絶対パス。拡張子を含む。
    '   Sampling    : サンプリング周波数。
    '   BPM         : 曲の速さ。
    '   ScoreType   : 音符の最小単位。
    '   Instrument  : 使用する楽器の名前。
    '
    '@Return
    '   なし。WAVファイルが出力される。
    '
    '@Description
    '   ピアノロールにあるノーツを実際の音にする関数。
    '
    '@Note
    '   ミックスダウン処理を大幅に改善。無駄な処理を省いた。
    '   一番上の緑のバーのところに連符、一番下の緑のバーのところにサステインペダル(S,E)
    '
    '----------------------------------------------------------------
        
    
    If PIANOROLL_END_COLUMN = -1 Then MsgBox ("曲の終わりが設定されていません。"): Exit Sub
    
    Dim BasicOneCellSeconds As Double                   '基本となる1セルあたりの秒数
    Dim OneCellSeconds As Double                        '1セルあたりの秒数(可変)
    Dim GenerateSeconds As Double                       '生成する音の長さ(サステイン対応用)
    Dim Is_Begin As Boolean, Is_End As Boolean          'ノーツの始まりと終わりを検知する
    Dim i As Long, j As Long, n As Long                 'カウンタ
    Dim MusicData() As Double                           '出力用波形データ
    Dim CurrentPos As Long                              '現在位置(配列のどの要素か)
    Dim InsertPos As Long                               '音の挿入位置(開始点)
    Dim Volume As Integer                               '音のボリュームの大きさ(%表示)
    Dim Is_Sustain As Boolean                           'サステイン状態かどうか
    Dim StartOfSustain As Long                          'サステインの始まりの列
    Dim EndOfSustain As Long: EndOfSustain = 0          'サステインの終わりの列
    Dim LongOfSustain As Double                         'サステインの長さ
    Const Release_Time As Double = 0.01                 'リリースタイム
    
    '初期化
    BasicOneCellSeconds = CSettings.Get_OneCellSeconds(BPM, ScoreType)
    
    ReDim MusicData(Get_Data_Count(Sampling, BPM, ScoreType) + Sampling * Release_Time) As Double
    For n = 0 To UBound(MusicData)
        MusicData(n) = 0    '0で初期化
    Next n
    
    '音源の読み込み
    Call Read_InstrumentSound(Instrument)
    
    For i = PIANOROLL_START_ROW To PIANOROLL_END_ROW
        
        '初期化
        CurrentPos = 0
        InsertPos = 0
        StartOfSustain = 0
        EndOfSustain = -2
        GenerateSeconds = 0
        Is_Begin = False: Is_End = False
        
        Is_Sustain = False
        
        For j = PIANOROLL_START_COLUMN To PIANOROLL_END_COLUMN
            With ThisWorkbook.Worksheets("ピアノロール")
                
                '連符があるか調べ、1セル当たりの秒数を変える
                If .Cells(NOTE_LENGTH_ROW, j).Value <> "" Then
                    OneCellSeconds = 2 * BasicOneCellSeconds / CInt(.Cells(NOTE_LENGTH_ROW, j).Value)
                Else
                    OneCellSeconds = BasicOneCellSeconds
                End If
                
                '始まりを検知
                If .Cells(i, j).Borders(xlEdgeLeft).LineStyle = xlContinuous And .Cells(i, j).Interior.Color = CSettings.Get_NoteColor() Then
                    Is_Begin = True '音の生成を開始
                    'サステインが設定されていた場合、開始点として再設定
                    If Is_Sustain = True Then
                        StartOfSustain = j
                    End If
                    
                    'ボリュームを設定
                    If .Cells(i, j).Value <> "" Then
                        Volume = CInt(.Cells(i, j).Value)
                    Else
                        Volume = 80
                    End If
                    
                    InsertPos = CurrentPos
                End If
                
                If Is_Begin = True Then
                    'ノーツが設定されている状態の時、生成する音の長さを更新
                    GenerateSeconds = GenerateSeconds + OneCellSeconds
                End If
                
                'サステインの始まりを検知
                If UCase(.Cells(SUSTAIN_ROW, j).Value) = "S" Then
                    
                    Is_Sustain = True   'サステイン有効にセット
                    StartOfSustain = j  'サステイン開始列をセット
                    EndOfSustain = -2   '終わりを未設定にセット
                    
                    'サステインの終わりを調べる
                    If EndOfSustain = -2 Then   '一番初めのみサステインの終わりを調べる
                        Dim SustainInfo() As Double: SustainInfo = Get_EndOfSustain(StartOfSustain, BasicOneCellSeconds)
                        EndOfSustain = CLng(SustainInfo(0))
                        LongOfSustain = SustainInfo(1)
                        'エラーチェック
                        If EndOfSustain = -1 Then
                            MsgBox ("サステインの終わりを設定してください。")
                            Exit Sub
                        ElseIf EndOfSustain > PIANOROLL_END_COLUMN Then
                            EndOfSustain = PIANOROLL_END_COLUMN
                        End If
                    End If
                End If
                
                
                If Is_Sustain = True And Is_Begin = True Then   'サステインが設定されていた場合
                    
                    GenerateSeconds = LongOfSustain
                    
                    LongOfSustain = LongOfSustain - OneCellSeconds
                    
                    'サステインが設定されていた場合には一気に終わりまで飛ぶ
                    Is_End = True
                    
                ElseIf Is_Sustain = True And Is_Begin = False Then  'サステインが設定されていて、ノーツはないとき
                    GenerateSeconds = 0
                    LongOfSustain = LongOfSustain - OneCellSeconds
                    
                ElseIf Is_Sustain = False And Is_Begin = True Then   'サステインが設定されていない場合
                    '終わりを検知
                    If .Cells(i, j).Borders(xlEdgeRight).LineStyle = xlContinuous And .Cells(i, j).Interior.Color = CSettings.Get_NoteColor() Then
                        Is_End = True
                    End If
                    
                ElseIf Is_Sustain = False And Is_Begin = False Then
                    GenerateSeconds = 0
                End If
                
                
                
                'サステインの終わりに来た時
                If Is_Sustain = True And j = EndOfSustain Then
                    LongOfSustain = 0
                    Is_Sustain = False
                End If
                
                
                
                '配列における現在位置を更新
                CurrentPos = CurrentPos + (Sampling * OneCellSeconds)
                
                '始まりと終わりが両方見つかったときのみ波形計算を行う
                If Is_Begin = True And Is_End = True Then
                    If Is_Sustain = False Then
                        Call Add_Sound(MusicData, Generate_Piano_Sound(i, Sampling, GenerateSeconds + Release_Time, Release_Time, Volume), InsertPos)
                    ElseIf Is_Sustain = True Then
                        Call Add_Sound(MusicData, Generate_Piano_Sound(i, Sampling, GenerateSeconds + 0.1, 0.1, Volume), InsertPos)
                    End If
                    '諸々リセット
                    GenerateSeconds = 0
                    Is_Begin = False: Is_End = False
                End If
                
                
            End With
        Next j
        
    Next i
    
    Dim MusicData2() As Double: MusicData2 = MusicData
    
    'Call CWav.Write_Wave_16bits_Mono(FilePath, Sampling, MusicData)
    Call CWav.Write_Wave_16bits_Stereo(FilePath, Sampling, MusicData, MusicData2)
    
    MsgBox ("WAVファイルへの書き出しが完了しました。")
    
End Sub


Public Function Get_EndOfSustain(ByVal StartOfSustain As Long, ByVal BasicOneCellSeconds As Double) As Double()
    Dim ret(1) As Double: ret(0) = -1: ret(1) = 0 '戻り値(終わりの列, 終わりまでの秒数)
    Dim OneCellSeconds As Double
    Dim i As Long: For i = StartOfSustain To PIANOROLL_END_COLUMN
        With ThisWorkbook.Worksheets("ピアノロール")
            '連符があるか調べ、1セル当たりの秒数を変える
            If .Cells(NOTE_LENGTH_ROW, i).Value <> "" Then
                OneCellSeconds = 2 * BasicOneCellSeconds / CInt(.Cells(NOTE_LENGTH_ROW, i).Value)
            Else
                OneCellSeconds = BasicOneCellSeconds
            End If
            
            ret(1) = ret(1) + OneCellSeconds
        
            If UCase(.Cells(SUSTAIN_ROW, i).Value) = "E" Then
                ret(0) = i
                Exit For
            End If
        End With
    Next i
    
    Get_EndOfSustain = ret
    
End Function

