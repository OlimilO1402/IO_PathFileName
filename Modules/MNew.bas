Attribute VB_Name = "MNew"
Option Explicit

Public Function PathFileName(ByVal aPathOrPFN As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathOrPFN, aFileName, aExt
End Function

'Public Function PFNStreamW(PFN As PathFileName) As PFNStreamW
'    Set PFNStreamW = New PFNStreamW: PFNStreamW.New_ PFN
'End Function

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

