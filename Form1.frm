VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Gif-X"
   ClientHeight    =   1695
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   5430
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "Hex Frame Locations"
      Top             =   120
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   120
      Pattern         =   "*.gif"
      TabIndex        =   0
      ToolTipText     =   "Click One To See"
      Top             =   120
      Width           =   3135
   End
   Begin VB.Menu mnuKill 
      Caption         =   "Kill All Gif Temp File"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub File1_Click()
    Dim TheFile As String, sFileName As String, FrameID, FLen
    
    CleanGifFrame
    sFileName = File1.Path & "\" & File1.List(File1.ListIndex)
    'MsgBox GetFileNameFromPath(sFileName, False)
    Open sFileName For Binary As #1
        FLen = FileLen(sFileName)
        TheFile = String(FLen, Chr(32))
        Get #1, 1, TheFile
    Close 1
    FrameID = Chr(33) & Chr(249)
    X = InStr(1, TheFile, FrameID)
    C = 1
    Do Until X = 0
        GifFrame(C) = X
        C = C + 1
        X = InStr(X + 1, TheFile, FrameID)
    Loop
    List1.Clear
    GifFrame(C) = FLen
    Dim TheFrame As String
    Dim TheHeader As String
    TheHeader = Left(TheFile, GifFrame(1) - 1)
    For i = 1 To C - 1
        List1.AddItem GifFrame(i)
        GifName = GetFileNameFromPath(sFileName, False)
        Open App.Path & "\" & GetFileNameFromPath(sFileName, False) & " (frame " & Format(i, "0#") & ").bmp" For Binary As #1
            'MsgBox TheFrame
            TheFrame = Mid(TheFile, GifFrame(i), Between(GifFrame(i), GifFrame(i + 1)) - 1)
            No = ""
            For D = 1 To 8
                Select Case D
                    Case 3, 5, 7
                        No = No & " " & Format(Hex(Asc(Mid(TheFrame, D, 1))), "00")
                    Case Else
                        No = No & Format(Hex(Asc(Mid(TheFrame, D, 1))), "00")
                End Select
            Next
            'MsgBox No, , "Frame #" & I
            Put #1, 1, TheHeader & TheFrame
        Close 1
    Next
    'Me.Caption = C - 1
    If Form2.Visible = True Then
        'Unload Form2
        Form2.Hide
        Y = 1
        Form2.Picture1.Picture = LoadPicture(App.Path & "\" & GifName & " (frame " & Format(Y, "0#") & ").bmp")
        With Form2
            .Width = Form2.Picture1.Width + 110
            .Height = Form2.Picture1.Height + 380
        End With
        Load Form2
        Form2.Show
    Else
        Load Form2
        Form2.Show
    End If
End Sub

Private Sub Form_Load()
    'File1.Path = "C:\Documents and Settings\christine\Desktop\Jeff\Pix\"
    File1.Path = App.Path & "\Gifs\"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub List1_Click()
    Dim TheFile As String, sFileName As String, FrameID
    'Dim TheFile2 As String, sFileName As String, FrameID
    sFileName = File1.Path & "\" & File1.List(File1.ListIndex)
    Open sFileName For Binary As #1
        TheFile = String(FileLen(sFileName), Chr(32))
        Get #1, List1.List(List1.ListIndex), TheFile
    Close 1
    FrameID = Chr(33) & Chr(249)
    X = InStr(1, TheFile, FrameID)
    C = 1
    Do Until X = 0
        GifFrame(C) = X
        C = C + 1
        X = InStr(X + 1, TheFile, FrameID)
    Loop
End Sub

Private Sub mnuKill_Click()
On Error GoTo 10
    Kill App.Path & "\*.bmp"
10
    Exit Sub
    MsgBox Err.Description, , "Error #" & Err.Number
End Sub
