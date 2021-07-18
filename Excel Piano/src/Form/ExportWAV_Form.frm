VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExportWAV_Form 
   Caption         =   "Excel Piano - WAV�t�@�C���o��"
   ClientHeight    =   1332
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7176
   OleObjectBlob   =   "ExportWAV_Form.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "ExportWAV_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CSound As New Sound
Private CSettings As New Settings

Private Sub Browse_Button_Click()
    Dim FileName As String
    Dim Title As String: Title = CSettings.Get_Title
    
    '�Ȗ������͂���Ă��Ȃ������ꍇ
    If Title = "" Then Title = "Music"
    
    
    FileName = CStr(Application.GetSaveAsFilename(InitialFileName:=Title, FileFilter:="WAV�t�@�C��,*.wav")): If FileName = "False" Then Exit Sub
    
    
    SaveTo_TextBox.Text = FileName
    
End Sub

Private Sub Cancel_Button_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    SaveTo_TextBox.Text = CSettings.Get_SavePath
    InstrumentName_Label.Caption = "�y��F" + CSettings.Get_UseInstrument
End Sub

Private Sub Write_Button_Click()
    If SaveTo_TextBox.Text = "" Then MsgBox ("�ۑ�����w�肵�Ă��������B"): Exit Sub
    
    '�o�͐��ۑ�����ꍇ�A�ݒ�ɏ�������
    If Save_SavePathCheckBox.Value = True Then
        CSettings.Save_SavePath (SaveTo_TextBox.Text)
    Else
        CSettings.Save_SavePath ("")
    End If
    
    '�Ȃ̏I�����擾
    PIANOROLL_END_COLUMN = CSettings.Get_LastColumn
    
    '�����o���J�n
    Call CSound.Export_Mixdown(SaveTo_TextBox.Text, 44100, CSettings.Get_Tempo, CSettings.Get_ScoreLength, CSettings.Get_UseInstrument)
    
End Sub

