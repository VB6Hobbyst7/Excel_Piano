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

'���̔g�`���v�Z����

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
    '   freqRow         : ���̍�����\�����g��������s�B
    '   Sampling        : �T���v�����O���g���B
    '   Seconds         : ���𐶐����鎞�Ԃ̒���(�b)�B
    '   Is_SoundBreak   : �����r�؂�邩�ǂ���
    '
    '@Return
    '   �v�Z���ꂽ�s�A�m���̔g�`��Double�^�̔z��ŕԂ����B
    '
    '@Description
    '   �C�ӂ̉��K�̃s�A�m���𐶐�����֐��B
    '
    '@Note
    '   �Ȃ��B
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
    '   freqRow         : ���̍�����\�����g��������s�B
    '   Sampling        : �T���v�����O���g���B
    '   Seconds         : ���𐶐����鎞�Ԃ̒���(�b)�B
    '   Release_Time    : �����[�X�^�C��
    '   Volue           : �{�����[��
    '
    '@Return
    '   �v�Z���ꂽ�s�A�m���̔g�`��Double�^�̔z��ŕԂ����B
    '
    '@Description
    '   �C�ӂ̉��K�̃s�A�m���𐶐�����֐��B
    '
    '@Note
    '   �Ȃ��B
    '
    '----------------------------------------------------------------
    
    
    '�w��b�������z�񂩂�v�f�𔲂����
    Dim Result() As Double: ReDim Result(Sampling * Seconds) As Double
    Dim i As Long
    For i = 0 To UBound(Result)
        Result(i) = Preset(freqRow - 2).WaveForm(i) * (Volume / 100)
    Next i
    
    Dim j, k As Long
    '�t�F�[�h������������
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
    '   WaveFormData1   : �d�˂���g�`�f�[�^�B
    '   WaveFormData2   : �d�˂�g�`�f�[�^�B
    '   Pos             : �d�ˍ��킹�J�n�_�B�z��̓Y�����̕������w��B
    '
    '@Return
    '   �Ȃ��B�����̔z��ɓ����Ă���g�`�f�[�^���X�V�����B
    '
    '@Description
    '   �����d�ˍ��킹�Ęa���ɂ��邽�߂̊֐��B
    '
    '@Note
    '   �Ȃ��B
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
    '   BPM         : �Ȃ̑����B
    '   ScoreType   : �����̍ŏ��P�ʁB
    '
    '@Return
    '   �g�`�f�[�^������f�[�^����Ԃ��B
    '
    '@Description
    '   �g�`�f�[�^������f�[�^����Ԃ��B
    '
    '@Note
    '   �Ȃ��B
    '
    '----------------------------------------------------------------
    
    Dim BasicOneCellSeconds As Double, OneCellSeconds As Double
    BasicOneCellSeconds = CSettings.Get_OneCellSeconds(BPM, ScoreType)
    
    
    Dim i As Long
    Dim Counter As Long: Counter = 0
    For i = PIANOROLL_START_COLUMN To PIANOROLL_END_COLUMN
        With ThisWorkbook.Worksheets("�s�A�m���[��")
            '�A�������邩���ׁA1�Z��������̕b����ݒ�
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
    '   FilePath    : �����o���ۑ���̃t�@�C���p�X�B��΃p�X�B�g���q���܂ށB
    '   Sampling    : �T���v�����O���g���B
    '   BPM         : �Ȃ̑����B
    '   ScoreType   : �����̍ŏ��P�ʁB
    '   Instrument  : �g�p����y��̖��O�B
    '
    '@Return
    '   �Ȃ��BWAV�t�@�C�����o�͂����B
    '
    '@Description
    '   �s�A�m���[���ɂ���m�[�c�����ۂ̉��ɂ���֐��B
    '
    '@Note
    '   �~�b�N�X�_�E��������啝�ɉ��P�B���ʂȏ������Ȃ����B
    '   ��ԏ�̗΂̃o�[�̂Ƃ���ɘA���A��ԉ��̗΂̃o�[�̂Ƃ���ɃT�X�e�C���y�_��(S,E)
    '
    '----------------------------------------------------------------
        
    
    If PIANOROLL_END_COLUMN = -1 Then MsgBox ("�Ȃ̏I��肪�ݒ肳��Ă��܂���B"): Exit Sub
    
    Dim BasicOneCellSeconds As Double                   '��{�ƂȂ�1�Z��������̕b��
    Dim OneCellSeconds As Double                        '1�Z��������̕b��(��)
    Dim GenerateSeconds As Double                       '�������鉹�̒���(�T�X�e�C���Ή��p)
    Dim Is_Begin As Boolean, Is_End As Boolean          '�m�[�c�̎n�܂�ƏI�������m����
    Dim i As Long, j As Long, n As Long                 '�J�E���^
    Dim MusicData() As Double                           '�o�͗p�g�`�f�[�^
    Dim CurrentPos As Long                              '���݈ʒu(�z��̂ǂ̗v�f��)
    Dim InsertPos As Long                               '���̑}���ʒu(�J�n�_)
    Dim Volume As Integer                               '���̃{�����[���̑傫��(%�\��)
    Dim Is_Sustain As Boolean                           '�T�X�e�C����Ԃ��ǂ���
    Dim StartOfSustain As Long                          '�T�X�e�C���̎n�܂�̗�
    Dim EndOfSustain As Long: EndOfSustain = 0          '�T�X�e�C���̏I���̗�
    Dim LongOfSustain As Double                         '�T�X�e�C���̒���
    Const Release_Time As Double = 0.01                 '�����[�X�^�C��
    
    '������
    BasicOneCellSeconds = CSettings.Get_OneCellSeconds(BPM, ScoreType)
    
    ReDim MusicData(Get_Data_Count(Sampling, BPM, ScoreType) + Sampling * Release_Time) As Double
    For n = 0 To UBound(MusicData)
        MusicData(n) = 0    '0�ŏ�����
    Next n
    
    '�����̓ǂݍ���
    Call Read_InstrumentSound(Instrument)
    
    For i = PIANOROLL_START_ROW To PIANOROLL_END_ROW
        
        '������
        CurrentPos = 0
        InsertPos = 0
        StartOfSustain = 0
        EndOfSustain = -2
        GenerateSeconds = 0
        Is_Begin = False: Is_End = False
        
        Is_Sustain = False
        
        For j = PIANOROLL_START_COLUMN To PIANOROLL_END_COLUMN
            With ThisWorkbook.Worksheets("�s�A�m���[��")
                
                '�A�������邩���ׁA1�Z��������̕b����ς���
                If .Cells(NOTE_LENGTH_ROW, j).Value <> "" Then
                    OneCellSeconds = 2 * BasicOneCellSeconds / CInt(.Cells(NOTE_LENGTH_ROW, j).Value)
                Else
                    OneCellSeconds = BasicOneCellSeconds
                End If
                
                '�n�܂�����m
                If .Cells(i, j).Borders(xlEdgeLeft).LineStyle = xlContinuous And .Cells(i, j).Interior.Color = CSettings.Get_NoteColor() Then
                    Is_Begin = True '���̐������J�n
                    '�T�X�e�C�����ݒ肳��Ă����ꍇ�A�J�n�_�Ƃ��čĐݒ�
                    If Is_Sustain = True Then
                        StartOfSustain = j
                    End If
                    
                    '�{�����[����ݒ�
                    If .Cells(i, j).Value <> "" Then
                        Volume = CInt(.Cells(i, j).Value)
                    Else
                        Volume = 80
                    End If
                    
                    InsertPos = CurrentPos
                End If
                
                If Is_Begin = True Then
                    '�m�[�c���ݒ肳��Ă����Ԃ̎��A�������鉹�̒������X�V
                    GenerateSeconds = GenerateSeconds + OneCellSeconds
                End If
                
                '�T�X�e�C���̎n�܂�����m
                If UCase(.Cells(SUSTAIN_ROW, j).Value) = "S" Then
                    
                    Is_Sustain = True   '�T�X�e�C���L���ɃZ�b�g
                    StartOfSustain = j  '�T�X�e�C���J�n����Z�b�g
                    EndOfSustain = -2   '�I���𖢐ݒ�ɃZ�b�g
                    
                    '�T�X�e�C���̏I���𒲂ׂ�
                    If EndOfSustain = -2 Then   '��ԏ��߂̂݃T�X�e�C���̏I���𒲂ׂ�
                        Dim SustainInfo() As Double: SustainInfo = Get_EndOfSustain(StartOfSustain, BasicOneCellSeconds)
                        EndOfSustain = CLng(SustainInfo(0))
                        LongOfSustain = SustainInfo(1)
                        '�G���[�`�F�b�N
                        If EndOfSustain = -1 Then
                            MsgBox ("�T�X�e�C���̏I����ݒ肵�Ă��������B")
                            Exit Sub
                        ElseIf EndOfSustain > PIANOROLL_END_COLUMN Then
                            EndOfSustain = PIANOROLL_END_COLUMN
                        End If
                    End If
                End If
                
                
                If Is_Sustain = True And Is_Begin = True Then   '�T�X�e�C�����ݒ肳��Ă����ꍇ
                    
                    GenerateSeconds = LongOfSustain
                    
                    LongOfSustain = LongOfSustain - OneCellSeconds
                    
                    '�T�X�e�C�����ݒ肳��Ă����ꍇ�ɂ͈�C�ɏI���܂Ŕ��
                    Is_End = True
                    
                ElseIf Is_Sustain = True And Is_Begin = False Then  '�T�X�e�C�����ݒ肳��Ă��āA�m�[�c�͂Ȃ��Ƃ�
                    GenerateSeconds = 0
                    LongOfSustain = LongOfSustain - OneCellSeconds
                    
                ElseIf Is_Sustain = False And Is_Begin = True Then   '�T�X�e�C�����ݒ肳��Ă��Ȃ��ꍇ
                    '�I�������m
                    If .Cells(i, j).Borders(xlEdgeRight).LineStyle = xlContinuous And .Cells(i, j).Interior.Color = CSettings.Get_NoteColor() Then
                        Is_End = True
                    End If
                    
                ElseIf Is_Sustain = False And Is_Begin = False Then
                    GenerateSeconds = 0
                End If
                
                
                
                '�T�X�e�C���̏I���ɗ�����
                If Is_Sustain = True And j = EndOfSustain Then
                    LongOfSustain = 0
                    Is_Sustain = False
                End If
                
                
                
                '�z��ɂ����錻�݈ʒu���X�V
                CurrentPos = CurrentPos + (Sampling * OneCellSeconds)
                
                '�n�܂�ƏI��肪�������������Ƃ��̂ݔg�`�v�Z���s��
                If Is_Begin = True And Is_End = True Then
                    If Is_Sustain = False Then
                        Call Add_Sound(MusicData, Generate_Piano_Sound(i, Sampling, GenerateSeconds + Release_Time, Release_Time, Volume), InsertPos)
                    ElseIf Is_Sustain = True Then
                        Call Add_Sound(MusicData, Generate_Piano_Sound(i, Sampling, GenerateSeconds + 0.1, 0.1, Volume), InsertPos)
                    End If
                    '���X���Z�b�g
                    GenerateSeconds = 0
                    Is_Begin = False: Is_End = False
                End If
                
                
            End With
        Next j
        
    Next i
    
    Dim MusicData2() As Double: MusicData2 = MusicData
    
    'Call CWav.Write_Wave_16bits_Mono(FilePath, Sampling, MusicData)
    Call CWav.Write_Wave_16bits_Stereo(FilePath, Sampling, MusicData, MusicData2)
    
    MsgBox ("WAV�t�@�C���ւ̏����o�����������܂����B")
    
End Sub


Public Function Get_EndOfSustain(ByVal StartOfSustain As Long, ByVal BasicOneCellSeconds As Double) As Double()
    Dim ret(1) As Double: ret(0) = -1: ret(1) = 0 '�߂�l(�I���̗�, �I���܂ł̕b��)
    Dim OneCellSeconds As Double
    Dim i As Long: For i = StartOfSustain To PIANOROLL_END_COLUMN
        With ThisWorkbook.Worksheets("�s�A�m���[��")
            '�A�������邩���ׁA1�Z��������̕b����ς���
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
