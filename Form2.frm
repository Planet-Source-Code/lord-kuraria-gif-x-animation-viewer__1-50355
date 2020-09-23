VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gif Viewer"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3240
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   125
      Left            =   1080
      Top             =   1440
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2130
      Left            =   0
      Picture         =   "Form2.frx":0000
      ScaleHeight     =   2130
      ScaleWidth      =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   1560
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Y

Private Sub Form_Load()
    If GifName <> "" Then
    Y = 1
    Picture1.Picture = LoadPicture(App.Path & "\" & GifName & " (frame " & Format(Y, "0#") & ").bmp")
    'With Me
    '    .Width = Picture1.Width + 110
    '    .Height = Picture1.Height + 380
    'End With
    Timer1.Enabled = True
    Me.Move 15, 15
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Timer1.Enabled = False
End Sub

Private Sub Picture1_Click()
    Timer1.Enabled = False
    Me.Hide
End Sub

Private Sub Timer1_Timer()
    'On Error GoTo 20
    'With Me
    '    .Width = Picture1.Width + 110
    '    .Height = Picture1.Height + 380
    'End With
    If Y < C - 1 Then
        Y = Y + 1
        'DoEvents
        Picture1.Picture = LoadPicture(App.Path & "\" & GifName & " (frame " & Format(Y, "0#") & ").bmp")
    Else
        Y = 1
        DoEvents
        Picture1.Picture = LoadPicture(App.Path & "\" & GifName & " (frame " & Format(Y, "0#") & ").bmp")
    End If
    DoEvents
    Exit Sub
20
    'MsgBox App.Path & "\" & GifName & " (frame " & Y & ").bmp" & vbNewLine & "C:\Documents and Settings\christine\Desktop\Poet of Prophecy Inc\Gif98\reader\dbz 14 (frame 1).bmp", , Err.Number & " - " & Err.Description
    DoEvents
    Unload Form2
    Err.Clear
End Sub
