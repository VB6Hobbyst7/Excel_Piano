VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Excel Piano - ���̓c�[�� "
   ClientHeight    =   3300
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8964.001
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Declare PtrSafe Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private CSettings As New Settings
Private CScoreData As New ScoreData
Private CMIDI As MIDI

Private OptBtn_Collection As Collection
Private TglBtn_Collection As Collection

Private Sub GoFirst_Button_Click()
    Dim Mode_Backup As String: Mode_Backup = Mode
    Mode = "Select"
    ThisWorkbook.Worksheets("�s�A�m���[��").Cells(PIANOROLL_START_ROW, PIANOROLL_START_COLUMN).Select
    Mode = Mode_Backup
End Sub

Private Sub InputMode_EndMarker_OptionButton_Click()
    NoteOrEnd = "End"
End Sub

Private Sub InputMode_Note_OptionButton_Click()
    NoteOrEnd = "Note"
End Sub

'���A���^�C�����t���~
Private Sub Stop_Button_Click()
    Me.Play_Button.Enabled = True
    Me.Stop_Button.Enabled = False
    Is_Playing = False
    CMIDI.End_Play
End Sub

'���A���^�C�����t���J�n
Private Sub Play_Button_Click()
    '�ݒ��ۑ�
    CSettings.Save_Settings_All
    Me.Play_Button.Enabled = False
    Me.Stop_Button.Enabled = True
    '�I���𒲂ׂ�
    PIANOROLL_END_COLUMN = CSettings.Get_LastColumn
    
    Set CMIDI = New MIDI
    CMIDI.Start (ActiveCell.Column)
End Sub


Private Sub GoLast_Button_Click()
    PIANOROLL_END_COLUMN = CSettings.Get_LastColumn
    If PIANOROLL_END_COLUMN = -1 Then Exit Sub
    Dim Mode_Backup As String: Mode_Backup = Mode
    Mode = "Select"
    ThisWorkbook.Worksheets("�s�A�m���[��").Cells(PIANOROLL_START_ROW, PIANOROLL_END_COLUMN).Select
    Mode = Mode_Backup
End Sub

Private Sub Import_EMD_Button_Click()
    
    Dim OpenFileName As Variant
    OpenFileName = Application.GetOpenFilename("Excel Music Data,*.emd")
    
    If OpenFileName <> False Then
        Call CScoreData.Import_ScoreData_EMD(OpenFileName)
        
        Tempo_TextBox.Text = CSettings.Get_Tempo                        '�e���|��ǂݍ���
        Call CSettings.Set_ComboBox_Rhythm(Rhythm_ComboBox)             '���q���R���{�{�b�N�X�ɓǂݍ���
        NoteColor_Label.BackColor = CSettings.Get_NoteColor             '�m�[�g�J���[��ǂݍ���
        EndColor_Label.BackColor = CSettings.Get_EndColor               '�G���h�}�[�J�[�J���[��ǂݍ���
        Title_TextBox.Text = CSettings.Get_Title                        '�Ȗ���ǂݍ���
    End If
    
End Sub

Private Sub Export_EMD_Button_Click()
    '�y�o�͌`��(�g���q��*.emd)�z
    '1�s�ڂɋȖ�
    '2�s�ڂɑ��x
    '3�s�ڂɍŏ��̉����̒P��
    '4�s�ڂɔ��q
    '���̂��ƁA�s�E�n�܂�̃Z���E�I���̃Z�����J���}��؂�ŏ��Ԃɓ����
    '
    '�y��z
    'God Knows...
    '150
    '16��
    '4/4
    '�I���̗�
    '22,6,9
    '(��)
    '
    '�Ƃ����悤�Ȍ`�ŃG�N�X�|�[�g����
    
    
    Dim FileName As String: FileName = CStr(Application.GetSaveAsFilename(InitialFileName:=Title_TextBox.Text, FileFilter:="EMD�t�@�C��,*.emd")): If FileName = "False" Then Exit Sub
    
    '����ۑ�
    Call CSettings.Save_Settings_All
    
    '�I���̗�𒲂ׂĂ���
    PIANOROLL_END_COLUMN = CSettings.Get_LastColumn
       
    '�o�͊J�n
    Call CScoreData.Export_ScoreData_EMD(FileName)
    
    MsgBox ("�y���̏����o�����������܂���")
End Sub

Private Sub Export_WAV_Button_Click()
    CSettings.Save_Settings_All
    ExportWAV_Form.Show vbModeless
End Sub

'�y���̍ŏ��P�ʂ��ύX���ꂽ�Ƃ�
Private Sub ScoreLength_ComboBox_Change()

    Select Case ScoreLength_ComboBox
        Case "�S"
            Threshold = 5
        Case "2��"
            Threshold = 4
        Case "4��"
            Threshold = 3
        Case "8��"
            Threshold = 2
        Case "16��"
            Threshold = 1
        Case "32��"
            Threshold = 0
        Case "64��"
            Threshold = -1
    End Select
    
    Dim Ctrl As Control
    For Each Ctrl In Me.Assist_Frame.Controls
        If Ctrl.Tag < Threshold Then
            Ctrl.Enabled = False
        Else
            Ctrl.Enabled = True
        End If
    Next
    
End Sub

Private Sub Close_Button_Click()
    Unload Me
End Sub

'�G���h�}�[�J�[�J���[��I������_�C�A���O��\�����Đݒ肷��{�^��
Private Sub Select_EndColor_Button_Click()
    ChooseColor_Form.Preview_Label.BackColor = CSettings.Get_EndColor
    ChooseColor_Form.Determine_Button.Tag = "End"
    ChooseColor_Form.Show
    EndColor_Label.BackColor = CSettings.Get_EndColor               '�G���h�}�[�J�[�J���[��ǂݍ���
    
End Sub

'�m�[�g�J���[��I������_�C�A���O��\�����Đݒ肷��{�^��
Private Sub Select_NoteColor_Button_Click()
    ChooseColor_Form.Preview_Label.BackColor = CSettings.Get_NoteColor
    ChooseColor_Form.Determine_Button.Tag = "Note"
    ChooseColor_Form.Show
    NoteColor_Label.BackColor = CSettings.Get_NoteColor             '�m�[�g�J���[��ǂݍ���
    
End Sub


'�R���g���[���C�x���g�̐ݒ�
Private Sub UserForm_Initialize()
    Dim Ctrl As Control
    Dim CCatchEvent As CatchEvent
    
    '���̓A�V�X�g�p�I�v�V�����{�^��
    Set OptBtn_Collection = New Collection
    For Each Ctrl In Me.Assist_Frame.Controls
        Set CCatchEvent = New CatchEvent
        CCatchEvent.SetCtrl Ctrl
        OptBtn_Collection.Add CCatchEvent
    Next
    
    '���̓��[�h��ݒ肷��g�O���{�^��
    Set TglBtn_Collection = New Collection
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "ToggleButton" Then
            Set CCatchEvent = New CatchEvent
            CCatchEvent.SetCtrl Ctrl
            TglBtn_Collection.Add CCatchEvent
        End If
    Next
    
    
    '***************************�ȉ��A�ǂݍ��ݎ��̏���*********************************
    Call CSettings.Set_ComboBox_ScoreLength(ScoreLength_ComboBox)   '�����̍ŏ��P�ʂ��R���{�{�b�N�X�ɓǂݍ���
    Tempo_TextBox.Text = CSettings.Get_Tempo                        '�e���|��ǂݍ���
    Call CSettings.Set_ComboBox_Rhythm(Rhythm_ComboBox)             '���q���R���{�{�b�N�X�ɓǂݍ���
    NoteColor_Label.BackColor = CSettings.Get_NoteColor             '�m�[�g�J���[��ǂݍ���
    EndColor_Label.BackColor = CSettings.Get_EndColor               '�G���h�}�[�J�[�J���[��ǂݍ���
    Title_TextBox.Text = CSettings.Get_Title                        '�Ȗ���ǂݍ���
    Call CSettings.Set_ComboBox_InstrumentList(Instrument_ComboBox)                  '�y���ǂݍ���
    Instrument_ComboBox.Text = CSettings.Get_UseInstrument
    
    'OptionButton�̃f�t�H���g��I��
    QuarterNoteOptionButton.Value = True
    InputMode_Note_OptionButton.Value = True
    
    
    '���̓��[�h��ݒ�
    Select_ToggleButton.BackColor = RGB(0, 255, 0)
    Select_ToggleButton.Value = True
    Mode = Select_ToggleButton.Tag
    
    'ExcelPiano�A�C�R����ǂݍ���
    Dim FileDir As String: Dim SoundDir As String: FileDir = Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\", InStrRev(ThisWorkbook.Path, "\") - 1) - 1) + "\Image\"
    Icon_Image.Picture = LoadPicture(FileDir + "Icon.jpg")
    
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Mode = "Select"
    Call CSettings.Save_Settings_All
    
    
    'MIDI�Đ������Ă������~
    If Is_Playing = True Then
        CMIDI.End_Play
    End If
End Sub

