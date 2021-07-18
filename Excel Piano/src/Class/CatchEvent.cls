VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CatchEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'�C�x���g�⑫�p�N���X

Private WithEvents Option_Button As MSForms.OptionButton    '���̓A�V�X�g�p�I�v�V�����{�^��
Attribute Option_Button.VB_VarHelpID = -1
Private WithEvents Toggle_Button As MSForms.ToggleButton    '���̓��[�h�ݒ�p�̃g�O���{�^��
Attribute Toggle_Button.VB_VarHelpID = -1

Private WithEvents Color_Label As MSForms.Label                        '�F�̑I�����x��
Attribute Color_Label.VB_VarHelpID = -1

Private CSettings As New Settings

Public Sub SetCtrl(ByRef Ctrl As MSForms.Control)
    Select Case TypeName(Ctrl)
        Case "OptionButton"
            Set Option_Button = Ctrl
        Case "ToggleButton"
            Set Toggle_Button = Ctrl
        Case "Label"
            Set Color_Label = Ctrl
    End Select
    
End Sub

'�I�v�V�����{�^��
Private Sub Option_Button_Click()
    AssistMode = Option_Button.Caption
End Sub


'�g�O���{�^��
Private Sub Toggle_Button_Click()
    
    If Toggle_Button.BackColor = RGB(0, 255, 0) Then Toggle_Button.Value = True
    If Toggle_Button.Value = False Then Exit Sub
    
    
    Const DefaultColor As Long = &H8000000F
    
    Dim Ctrl As Control
    For Each Ctrl In MainForm.Controls
        If TypeName(Ctrl) = "ToggleButton" Then
            
            If Ctrl.Name = Toggle_Button.Name Then
                Ctrl.BackColor = RGB(0, 255, 0)
                Ctrl.Value = True
                Mode = Ctrl.Tag
            Else
                Ctrl.BackColor = DefaultColor
                Ctrl.Value = False
            End If
            
            
        End If
    Next
    
End Sub

Private Sub Color_Label_Click()
    ChooseColor_Form.Preview_Label.BackColor = Color_Label.BackColor
End Sub
