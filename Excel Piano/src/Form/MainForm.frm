VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MainForm 
   Caption         =   "Excel Piano - 入力ツール "
   ClientHeight    =   3300
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8964.001
   OleObjectBlob   =   "MainForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
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
    ThisWorkbook.Worksheets("ピアノロール").Cells(PIANOROLL_START_ROW, PIANOROLL_START_COLUMN).Select
    Mode = Mode_Backup
End Sub

Private Sub InputMode_EndMarker_OptionButton_Click()
    NoteOrEnd = "End"
End Sub

Private Sub InputMode_Note_OptionButton_Click()
    NoteOrEnd = "Note"
End Sub

'リアルタイム演奏を停止
Private Sub Stop_Button_Click()
    Me.Play_Button.Enabled = True
    Me.Stop_Button.Enabled = False
    Is_Playing = False
    CMIDI.End_Play
End Sub

'リアルタイム演奏を開始
Private Sub Play_Button_Click()
    '設定を保存
    CSettings.Save_Settings_All
    Me.Play_Button.Enabled = False
    Me.Stop_Button.Enabled = True
    '終わりを調べる
    PIANOROLL_END_COLUMN = CSettings.Get_LastColumn
    
    Set CMIDI = New MIDI
    CMIDI.Start (ActiveCell.Column)
End Sub


Private Sub GoLast_Button_Click()
    PIANOROLL_END_COLUMN = CSettings.Get_LastColumn
    If PIANOROLL_END_COLUMN = -1 Then Exit Sub
    Dim Mode_Backup As String: Mode_Backup = Mode
    Mode = "Select"
    ThisWorkbook.Worksheets("ピアノロール").Cells(PIANOROLL_START_ROW, PIANOROLL_END_COLUMN).Select
    Mode = Mode_Backup
End Sub

Private Sub Import_EMD_Button_Click()
    
    Dim OpenFileName As Variant
    OpenFileName = Application.GetOpenFilename("Excel Music Data,*.emd")
    
    If OpenFileName <> False Then
        Call CScoreData.Import_ScoreData_EMD(OpenFileName)
        
        Tempo_TextBox.Text = CSettings.Get_Tempo                        'テンポを読み込む
        Call CSettings.Set_ComboBox_Rhythm(Rhythm_ComboBox)             '拍子をコンボボックスに読み込む
        NoteColor_Label.BackColor = CSettings.Get_NoteColor             'ノートカラーを読み込む
        EndColor_Label.BackColor = CSettings.Get_EndColor               'エンドマーカーカラーを読み込む
        Title_TextBox.Text = CSettings.Get_Title                        '曲名を読み込む
    End If
    
End Sub

Private Sub Export_EMD_Button_Click()
    '【出力形式(拡張子は*.emd)】
    '1行目に曲名
    '2行目に速度
    '3行目に最小の音符の単位
    '4行目に拍子
    'そのあと、行・始まりのセル・終わりのセルをカンマ区切りで順番に入れる
    '
    '【例】
    'God Knows...
    '150
    '16分
    '4/4
    '終わりの列
    '22,6,9
    '(略)
    '
    'というような形でエクスポートする
    
    
    Dim FileName As String: FileName = CStr(Application.GetSaveAsFilename(InitialFileName:=Title_TextBox.Text, FileFilter:="EMDファイル,*.emd")): If FileName = "False" Then Exit Sub
    
    '情報を保存
    Call CSettings.Save_Settings_All
    
    '終わりの列を調べておく
    PIANOROLL_END_COLUMN = CSettings.Get_LastColumn
       
    '出力開始
    Call CScoreData.Export_ScoreData_EMD(FileName)
    
    MsgBox ("楽譜の書き出しが完了しました")
End Sub

Private Sub Export_WAV_Button_Click()
    CSettings.Save_Settings_All
    ExportWAV_Form.Show vbModeless
End Sub

'楽譜の最小単位が変更されたとき
Private Sub ScoreLength_ComboBox_Change()

    Select Case ScoreLength_ComboBox
        Case "全"
            Threshold = 5
        Case "2分"
            Threshold = 4
        Case "4分"
            Threshold = 3
        Case "8分"
            Threshold = 2
        Case "16分"
            Threshold = 1
        Case "32分"
            Threshold = 0
        Case "64分"
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

'エンドマーカーカラーを選択するダイアログを表示して設定するボタン
Private Sub Select_EndColor_Button_Click()
    ChooseColor_Form.Preview_Label.BackColor = CSettings.Get_EndColor
    ChooseColor_Form.Determine_Button.Tag = "End"
    ChooseColor_Form.Show
    EndColor_Label.BackColor = CSettings.Get_EndColor               'エンドマーカーカラーを読み込む
    
End Sub

'ノートカラーを選択するダイアログを表示して設定するボタン
Private Sub Select_NoteColor_Button_Click()
    ChooseColor_Form.Preview_Label.BackColor = CSettings.Get_NoteColor
    ChooseColor_Form.Determine_Button.Tag = "Note"
    ChooseColor_Form.Show
    NoteColor_Label.BackColor = CSettings.Get_NoteColor             'ノートカラーを読み込む
    
End Sub


'コントロールイベントの設定
Private Sub UserForm_Initialize()
    Dim Ctrl As Control
    Dim CCatchEvent As CatchEvent
    
    '入力アシスト用オプションボタン
    Set OptBtn_Collection = New Collection
    For Each Ctrl In Me.Assist_Frame.Controls
        Set CCatchEvent = New CatchEvent
        CCatchEvent.SetCtrl Ctrl
        OptBtn_Collection.Add CCatchEvent
    Next
    
    '入力モードを設定するトグルボタン
    Set TglBtn_Collection = New Collection
    For Each Ctrl In Me.Controls
        If TypeName(Ctrl) = "ToggleButton" Then
            Set CCatchEvent = New CatchEvent
            CCatchEvent.SetCtrl Ctrl
            TglBtn_Collection.Add CCatchEvent
        End If
    Next
    
    
    '***************************以下、読み込み時の処理*********************************
    Call CSettings.Set_ComboBox_ScoreLength(ScoreLength_ComboBox)   '音符の最小単位をコンボボックスに読み込む
    Tempo_TextBox.Text = CSettings.Get_Tempo                        'テンポを読み込む
    Call CSettings.Set_ComboBox_Rhythm(Rhythm_ComboBox)             '拍子をコンボボックスに読み込む
    NoteColor_Label.BackColor = CSettings.Get_NoteColor             'ノートカラーを読み込む
    EndColor_Label.BackColor = CSettings.Get_EndColor               'エンドマーカーカラーを読み込む
    Title_TextBox.Text = CSettings.Get_Title                        '曲名を読み込む
    Call CSettings.Set_ComboBox_InstrumentList(Instrument_ComboBox)                  '楽器を読み込む
    Instrument_ComboBox.Text = CSettings.Get_UseInstrument
    
    'OptionButtonのデフォルトを選択
    QuarterNoteOptionButton.Value = True
    InputMode_Note_OptionButton.Value = True
    
    
    '入力モードを設定
    Select_ToggleButton.BackColor = RGB(0, 255, 0)
    Select_ToggleButton.Value = True
    Mode = Select_ToggleButton.Tag
    
    'ExcelPianoアイコンを読み込み
    Dim FileDir As String: Dim SoundDir As String: FileDir = Left(ThisWorkbook.Path, InStrRev(ThisWorkbook.Path, "\", InStrRev(ThisWorkbook.Path, "\") - 1) - 1) + "\Image\"
    Icon_Image.Picture = LoadPicture(FileDir + "Icon.jpg")
    
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Mode = "Select"
    Call CSettings.Save_Settings_All
    
    
    'MIDI再生をしていたら停止
    If Is_Playing = True Then
        CMIDI.End_Play
    End If
End Sub

