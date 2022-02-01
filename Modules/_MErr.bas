Attribute VB_Name = "MErr"
Option Explicit
Public ErrLog As String

Public Function MessError(ClsName As String, FncName As String, _
                            Optional AddInfo As String = "", _
                            Optional bLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKOnly Or vbCritical) As VbMsgBoxResult
    If bLoud Then
        Dim sErr As String
        sErr = "Fehler: " & Err.Number & vbCrLf & _
               "in:     " & ClsName & "::" & FncName & vbCrLf & _
               "Info:   " & Err.Description & vbCrLf & AddInfo
        MsgBox sErr, vbDecor
        Err.Clear
    End If
    If bErrLog Then
        ErrLog = ErrLog & vbCrLf & Now & " " & sErr
    End If
End Function

Public Function MessErrorRetry(ClsName As String, FncName As String, _
                               Optional AddInfo As String = "", _
                               Optional bErrLog As Boolean = True) As VbMsgBoxResult
    MessErrorRetry = MessError(ClsName, FncName, AddInfo, True, bErrLog, vbRetryCancel)
End Function
