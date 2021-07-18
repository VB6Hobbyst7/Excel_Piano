VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ChooseColor_Form 
   Caption         =   "色の選択"
   ClientHeight    =   4380
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4104
   OleObjectBlob   =   "ChooseColor_Form.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ChooseColor_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CSettings As New Settings

Private Ctrl(1 To 56) As New CatchEvent

Private Sub Cancel_Button_Click()
    Unload Me
End Sub

Private Sub Determine_Button_Click()
    '色を保存
    If Determine_Button.Tag = "Note" Then
        Call CSettings.Save_NoteColor(Preview_Label.BackColor)
    Else
        Call CSettings.Save_EndColor(Preview_Label.BackColor)
    End If
    
    Unload Me
    
End Sub

Private Sub UserForm_Initialize()
    Dim r As Integer: r = 6 '横
    Dim c As Integer: c = 9 '縦
    Dim i As Integer, j As Integer  'カウンタ
    Dim ColorCounter As Integer
    
    Dim newLabel As MSForms.Label
    For i = 1 To r
        For j = 1 To c
            ColorCounter = ColorCounter + 1
            Set newLabel = Me.Controls.Add("Forms.Label.1", CStr(i) + CStr(j))
            With newLabel
                .Caption = "C"
                .Font.size = 30
                .AutoSize = True
                .Left = 10 + (j - 1) * .Width
                .Top = 10 + (i - 1) * .Height
                .BackColor = ThisWorkbook.Colors(ColorCounter)
                .ForeColor = ThisWorkbook.Colors(ColorCounter)
                .BorderStyle = fmBorderStyleSingle
            End With
            Ctrl(ColorCounter).SetCtrl newLabel
        Next j
    Next i
    
End Sub

