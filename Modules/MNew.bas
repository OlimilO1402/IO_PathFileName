Attribute VB_Name = "MNew"
Option Explicit

Public Function PathFileName(ByVal aPathOrPFN As String, _
                    Optional ByVal aFileName As String, _
                    Optional ByVal aExt As String) As PathFileName
    Set PathFileName = New PathFileName: PathFileName.New_ aPathOrPFN, aFileName, aExt
End Function

Public Function RecursiveReplace(ByVal Expression As String, ByVal Find As String, ByVal Replace As String) As String
    Dim pos As Long: pos = InStr(1, Expression, Find)
    If pos Then
        Dim r As String: r = VBA.Replace(Expression, Find, Replace)
        'check for stack overflow:
        If (r = Expression) Or (Len(Expression) < Len(r)) Then RecursiveReplace = r: Exit Function
        RecursiveReplace = RecursiveReplace(r, Find, Replace)
    Else
        RecursiveReplace = Expression
    End If
End Function

Public Function RecursiveReplaceSL(ByVal Expression As String, ByVal Find As String, ByVal Replace As String, Optional ByVal Start As Long = 1, Optional ByVal Length As Long = -1) As String
    'check input parameters return early if necessary
    If Length < 0 And Start = 1 Then RecursiveReplaceSL = RecursiveReplace(Expression, Find, Replace): Exit Function
    Dim le As Long: le = Len(Expression)
    If Start < 1 Or le < Start Then Exit Function 'return nothing
    If Length < 1 Or le < Start + Length Then Length = le - Start + 1
    
    'for debugging:
    Dim sl As String: sl = Left$(Expression, Start - 1)
    Dim sr As String: sr = Mid$(Expression, Length + 1)
    Expression = RecursiveReplace(Mid$(Expression, Start, Length), Find, Replace)
    RecursiveReplaceSL = sl & Expression & sr
    'same but shorter and less noise:
    'RecursiveReplaceSL = Left$(Expression, Start - 1) & RecursiveReplace(Mid$(Expression, Start, Length)) & Mid$(Expression, Start, Length)
End Function

