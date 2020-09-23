Attribute VB_Name = "Module1"

Global GifFrame(1 To 50) As Long 'As Collection
Global C As Long
Global GifName As String

Public Sub CleanGifFrame()
    For i = 1 To 50
        GifFrame(i) = 0
    Next
End Sub

Public Function Between(lFrom As Long, lTo As Long)
    For J = lFrom To lTo
        L = L + 1
    Next
    Between = L
End Function

Public Function GetFileNameFromPath(ByVal FullFilePath As String, bExtension As Boolean) As String
    Dim sFileName As String, FindSlash As Long
    
    If FullFilePath = "" Then Exit Function
    FindSlash = InStrRev(FullFilePath, "\")
    sFileName = Mid(FullFilePath, FindSlash + 1, Between(FindSlash, Len(FullFilePath)))
    If bExtension = True Then
        GetFileNameFromPath = sFileName
    Else
        GetFileNameFromPath = Left(sFileName, Len(sFileName) - Between(InStrRev(sFileName, "."), Len(sFileName)))
    End If
End Function
