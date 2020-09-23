Attribute VB_Name = "Module1"
Option Explicit

Public sTips() As String
Function GetTipTitle(ByVal sIn As String) As String
Dim sTMP() As String
    sTMP = Split(sIn, vbCrLf)
    GetTipTitle = sTMP(0)
End Function
Function FileExists(FileName As String) As Boolean
    On Error GoTo ErrorHandler
    FileExists = (GetAttr(FileName) And vbDirectory) = 0
ErrorHandler:
End Function
Function FileText(ByVal FileName As String) As String
    Dim handle As Integer
    
    If Len(Dir$(FileName)) = 0 Then
        Err.Raise 53
    End If
    
    handle = FreeFile
    Open FileName$ For Binary As #handle
    FileText = Space$(LOF(handle))
    Get #handle, , FileText
    Close #handle
End Function
