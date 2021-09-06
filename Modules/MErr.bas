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
'    If Err Then
'        MessErrorRetry = MsgBox("Fehler: " & Err.Number & vbCrLf & _
'                                "in:     " & ClsName & "::" & FncName & vbCrLf & _
'                                "Info:   " & Err.Description & vbCrLf & _
'                                AddInfo, vbRetryCancel)
'    End If
End Function

'Function IsInIDE() As Boolean
'Try: On Error GoTo Catch
'    Debug.Print 1 / 0
'    IsInIDE = False: Exit Function
'Catch: IsInIDE = True
'End Function
Function IsInIDE() As Boolean
    IsInIDE = App.LogMode = 0 'LogModeConstants.vbLogAuto
End Function

Function LogModeConstants_ToStr(e As LogModeConstants) As String
    Dim s As String
    Select Case e
    Case LogModeConstants.vbLogAuto:      s = "vbLogAuto"      ' 0
    Case LogModeConstants.vbLogOff:       s = "vbLogOff"       ' 1
    Case LogModeConstants.vbLogToFile:    s = "vbLogToFile"    ' 2
    Case LogModeConstants.vbLogToNT:      s = "vbLogToNT"      ' 3
    Case LogModeConstants.vbLogOverwrite: s = "vbLogOverwrite" '16 &H10
    Case LogModeConstants.vbLogThreadID:  s = "vbLogThreadID"  '32 &H20
    Case Else: s = CStr(e)
    End Select
    LogModeConstants_ToStr = s
End Function
