VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressBar_Form 
   Caption         =   "檢核表前處理進度"
   ClientHeight    =   1188
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5460
   OleObjectBlob   =   "ProgressBar_Form.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressBar_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'
Public Sub Init_Progress_Bar()
    With ProgressBar_Form
        .Progress_Frame.Width = 0
        .Progress_Label.Caption = "完成進度 0 %"
        .Show vbModeless
    End With
End Sub
'
'
'
Public Sub Set_Progress_Percentage(completePercentage As Double)
    Dim progressBarWidth    As Long
    Dim progressCaption     As String
    
    progressCaption = "完成進度 " & CStr(Round(completePercentage * 100, 0)) & " %"
    progressBarWidth = Round(completePercentage * ProgressBar_Form.Border_Frame.Width, 0)
        
    ProgressBar_Form.Progress_Label.Caption = progressCaption
    ProgressBar_Form.Progress_Frame.Width = progressBarWidth

End Sub
