VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'WAVファイルを扱うためのクラス

'--------------------------------------------------------------------------------------------
'Waveファイルの構造
'<RIFFChunk> + <SubChunk> + <SubChunk>
'
'SubChunkにはいくつか種類がある(fmt , data, fact, LIST)
'fmt + dataの場合にのみ対応
'fact, LISTについてはパス
'--------------------------------------------------------------------------------------------
'<RIFF>
'ChunkID        4bytes  "RIFF"
'ChunkSize      4bytes  これ以降のファイルサイズ (byte単位)
'Format         4bytes  "WAVE"
'--------------------------------------------------------------------------------------------
'<fmt>
'SubchunkID     4bytes  "fmt ", スペースありの4文字
'SubchunkSize   4bytes  SubChunkDataのサイズ, byte単位, "fmt "なら16
'AudioFormat    2bytes  フォーマットID, リニアPCMなら1
'NumChannels    2bytes  チャンネル数, モノラルは1, ステレオは2
'SampleRate     4bytes  サンプリングレート, Hz単位, 8kHz, 44.1kHz, etc
'ByteRate       4bytes  データ速度, bytes/sec単位, 44.1kHz ステレオ, 16bitsなら176400
'BlockAlign     2bytes  ブロックサイズ, bytes/sample単位, ステレオ, 16bitsなら4
'BitsPerSample  2bytes  サンプルサイズ, bits/sample単位, WAVEフォーマットでは8bits or 16bits
'--------------------------------------------------------------------------------------------
'<data>
'SubchunkID     4bytes  "data"
'SubchunkSize   4bytes    SubChunkDataのサイズ, byte単位
'data
'ステレオの場合、LRLR・・・の順
'8bitsなら符号なし(0～255、128が無音)
'16bitsなら符号つき (-32768～32767、0が無音)
'--------------------------------------------------------------------------------------------

'<RIFF>
Private Type RIFF_CHUNK
    ChunkID As String * 4
    ChunkSize As Long
    Format As String * 4
End Type

'<fmt>
Private Type FMT_CHUNK
    SubchunkID As String * 4
    SubchunkSize As Long
    AudioFormat As Integer
    NumChannels As Integer
    SampleRate As Long
    ByteRate As Long
    BlockAlign As Integer
    BitsPerSample As Integer
End Type

'<Data>
Private Type DATA_CHUNK
    SubchunkID As String * 4
    SubchunkSize As Long
End Type


Public Sub Write_Wave_16bits_Mono(ByVal FilePath As String, ByVal Sampling As Long, ByRef Data() As Double)
    '----------------------------------------------------------------
    '@Name
    '   Write_Wave_16bits_Mono
    '
    '@Param
    '   FilePath    : 保存先のファイルパス。絶対パス。拡張子含む。
    '   Sampling    : サンプリング周波数。
    '   Data()      : 波形データ。配列で入力される。
    '
    '@Return
    '   なし。
    '
    '@Description
    '   リニアPCM, モノラル, 16bitでWAVファイルに波形を書き出す。
    '
    '@Note
    '   リニアPCM, モノラル, 16bit以外は全く動かない。
    '
    '----------------------------------------------------------------
    
    Dim Length As Long: Length = UBound(Data)
    
    Dim RIFFChunk As RIFF_CHUNK
    Dim fmtChunk As FMT_CHUNK
    Dim DataChunk As DATA_CHUNK
    
    
    RIFFChunk.ChunkID = "RIFF"
    RIFFChunk.ChunkSize = Length * 2 + 36
    RIFFChunk.Format = "WAVE"
   
    fmtChunk.SubchunkID = "fmt "
    fmtChunk.SubchunkSize = 16
    fmtChunk.AudioFormat = 1
    fmtChunk.NumChannels = 1
    fmtChunk.SampleRate = Sampling
    fmtChunk.ByteRate = Sampling * 2
    fmtChunk.BlockAlign = 2
    fmtChunk.BitsPerSample = 16
   
    DataChunk.SubchunkID = "data"
    DataChunk.SubchunkSize = Length * 2
    
    'クリッピング
    Dim Output_Data() As Integer: ReDim Output_Data(UBound(Data)) As Integer
    Dim i As Long
    Dim tmp As Double
    For i = 0 To UBound(Data)
        tmp = (Data(i) + 1) / 2 * 65536
        If tmp > 65535 Then
            tmp = 65535
        ElseIf tmp < 0 Then
            tmp = 0
        End If
        Output_Data(i) = CInt((tmp - 0.5) - 32768)
    Next i
    
    
    Open FilePath For Binary As 1
        Put 1, , RIFFChunk
        Put 1, , fmtChunk
        Put 1, , DataChunk
        Put 1, , Output_Data
    Close 1
    
End Sub

'WAVEファイルを読み込み
Public Sub Read_Wave_16bits_Mono(ByVal WaveFilePath As String, ByRef WaveFormData() As Integer)
    '----------------------------------------------------------------
    '@Name
    '   Write_Wave_16bits_Mono
    '
    '@Param
    '   WaveFilePath        : 読み込むファイルのパス。
    '   WaveFormData()      : 波形データを入れておく変数。配列で入力される。
    '   Seconds             : 読み込む秒数。
    '
    '@Return
    '   なし。引数として設定された配列にデータが直接入る。
    '
    '@Description
    '   リニアPCM, モノラル, 16bitでWAVファイルに波形を読みだす。
    '
    '@Note
    '   リニアPCM, モノラル, 16bit以外は全く動かない。
    '
    '----------------------------------------------------------------
    
    Dim RIFFChunk As RIFF_CHUNK
    Dim fmtChunk As FMT_CHUNK
    Dim DataChunk As DATA_CHUNK

    
    Dim Sampling As Long
   
    Open WaveFilePath For Binary As #1  'WAVファイルをバイナリ形式で開く
        Get #1, , RIFFChunk 'RIFF識別子の読み込み
        Get #1, , fmtChunk  'fmt識別子の読み込み
        Get #1, , DataChunk '波形データに関する情報の読み込み
        ReDim WaveFormData(DataChunk.SubchunkSize) As Integer
        Get #1, , WaveFormData      '波形データの読み込み
        
    Close #1
    
End Sub




Public Sub Write_Wave_16bits_Stereo(ByVal FilePath As String, ByVal Sampling As Long, ByRef Data_Right() As Double, ByRef Data_Left() As Double)
      '----------------------------------------------------------------
    '@Name
    '   Write_Wave_16bits_Stereo
    '
    '@Param
    '   FilePath    : 保存先のファイルパス。絶対パス。拡張子含む。
    '   Sampling    : サンプリング周波数。
    '   DataRight() : 波形データ右。配列で入力される。
    '   Data_Left() : 波形データ左。配列で入力される。
    '
    '@Return
    '   なし。
    '
    '@Description
    '   リニアPCM, ステレオ, 16bitでWAVファイルに波形を書き出す。
    '
    '@Note
    '   リニアPCM, ステレオ, 16bit以外は全く動かない。
    '
    '----------------------------------------------------------------
    
    Dim Length As Long: Length = UBound(Data_Right) + UBound(Data_Left)
    
    Dim RIFFChunk As RIFF_CHUNK
    Dim fmtChunk As FMT_CHUNK
    Dim DataChunk As DATA_CHUNK
    
    
    RIFFChunk.ChunkID = "RIFF"
    RIFFChunk.ChunkSize = Length * 2 + 36
    RIFFChunk.Format = "WAVE"
   
    fmtChunk.SubchunkID = "fmt "
    fmtChunk.SubchunkSize = 16
    fmtChunk.AudioFormat = 1
    fmtChunk.NumChannels = 2
    fmtChunk.SampleRate = Sampling
    fmtChunk.ByteRate = Sampling * 2 * 2
    fmtChunk.BlockAlign = 4
    fmtChunk.BitsPerSample = 16
   
    DataChunk.SubchunkID = "data"
    DataChunk.SubchunkSize = Length * 2
    
    'クリッピング
    Dim Clipping_Data_Right() As Integer: ReDim Clipping_Data_Right(UBound(Data_Right)) As Integer
    Dim i As Long
    Dim tmp_Right As Double
    For i = 0 To UBound(Data_Right)
        tmp_Right = (Data_Right(i) + 1) / 2 * 65536
        If tmp_Right > 65535 Then
            tmp_Right = 65535
        ElseIf tmp_Right < 0 Then
            tmp_Right = 0
        End If
        Clipping_Data_Right(i) = CInt((tmp_Right - 0.5) - 32768)
    Next i
    
    
    Dim Clipping_Data_Left() As Integer: ReDim Clipping_Data_Left(UBound(Data_Left)) As Integer
    Dim j As Long
    Dim tmp_Left As Double
    For j = 0 To UBound(Data_Left)
        tmp_Left = (Data_Left(j) + 1) / 2 * 65536
        If tmp_Left > 65535 Then
            tmp_Left = 65535
        ElseIf tmp_Left < 0 Then
            tmp_Left = 0
        End If
        Clipping_Data_Left(j) = CInt((tmp_Left - 0.5) - 32768)
    Next j
    
    
    Dim Output_Data() As Integer: ReDim Output_Data(Length) As Integer
    Dim n As Long: For n = 0 To Length
        If n Mod 2 = 0 Then '偶数の時は左チャンネル
            Output_Data(n) = Clipping_Data_Right(n / 2)
        Else                '奇数の時は右チャンネル
            Output_Data(n) = Clipping_Data_Left((n - 1) / 2)
        End If
    Next n
    
    Open FilePath For Binary As 1
        Put 1, , RIFFChunk
        Put 1, , fmtChunk
        Put 1, , DataChunk
        Put 1, , Output_Data
    Close 1
End Sub


