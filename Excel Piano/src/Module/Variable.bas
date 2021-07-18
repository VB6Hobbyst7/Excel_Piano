Attribute VB_Name = "Variable"
Option Explicit

'�O���[�o���ϐ�

Public Const PIANOROLL_START_ROW As Integer = 2         '���Ղ̎n�܂�̍s
Public Const PIANOROLL_END_ROW As Integer = 89          '���Ղ̏I���̍s
Public Const PIANOROLL_START_COLUMN As Integer = 2      '���Ղ̎n�܂�̗�
Public PIANOROLL_END_COLUMN As Integer                  '���Ղ̏I���̗�
Public Const NOTE_LENGTH_ROW As Integer = 1             '�A���p�̗�
Public Const SUSTAIN_ROW As Integer = 90                '�T�X�e�C���p�̗�
Public Const KEYBOARD_COLUMN As Integer = 1             '���Ղ��`����Ă����

Public Mode As String           '���̓��[�h(�I���A�ǉ��A�폜�A�����A����)

Public AssistMode As String     '�A�V�X�g���[�h

Public Threshold As Integer     '���̓A�V�X�g�̖������𔻒f����臒l

Public NoteOrEnd As String      '�m�[�g���G���h�}�[�J�[�����͂���ۂɎg��

Public Is_Playing As Boolean    '���A���^�C�����t�����ǂ���(���t����True)

Private CSettings As New Settings

'�S���̃m�[�c��80������
Sub Set_Volume()
    Dim i As Long, j As Long
    With ThisWorkbook.Worksheets("�s�A�m���[��")
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

'���������ɃT�X�e�C����ݒ�ł��邩��?
Sub SetSustain()
    
    Const Interval As Integer = 8   '4/4 �ŏ�16�������̎���1���߂��ƂɃT�X�e�C��������
    
    Dim i As Long
    Dim OneCellSeconds As Double, BasicOneCellSeconds As Double
    BasicOneCellSeconds = CSettings.Get_OneCellSeconds(CSettings.Get_Tempo, CSettings.Get_ScoreLength)
    Dim StartColumn As Long: StartColumn = 6
    
    Dim tmp As Double
    For i = StartColumn To CSettings.Get_LastColumn
    
        '�A�������邩���ׁA1�Z��������̕b����ς���
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


