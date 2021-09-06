VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PathFileName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''1'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''2'''''''''''''''''''''''''
'''''''''1'''''''''2'''''''''3'''''''''4'''''''''5'''''''''6'''''''''7'''''''''8'''''''''9'''''''''0'''''''''1'''''''''2'''''''''3'''''''''4'''''''''5'''''''''6'''''''''7'''''''''8'''''''''9'''''''''0'''''''''1'''''''''2'''''
''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0''''5''''0'''''
'        1         2         3         4         5         6         7         8         9        10        11        12        13        14        15        16        17        18        19        20        21        22
'23456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345678901234567890123456789012345
'
'Class PathFileName aka PFN _P_ath_F_ile_N_ame
'works with:
' * VB6 (aka VBC aka VB-Classic)
' * VBA6 (VBA before 2007)
' * VBA7-x86
' * VBA7-X64
'
'a PathFileName consists of the following parts:
'Drive     |      Only path      |   Only filename  |  extension
'  C:\     |  Windows\System32\  |      user32      |    .dll
'
'resp:
'      Path (=Drive+OnlyPath)    |  Filename
'  C:\Windows\System32\          |  user32.dll

Option Explicit 'OM 06.11.2015 13:45; Zeilen: 336
#If VBA7 = 0 Then
    'Public Enum LongPtr
    '    [_]
    'End Enum
#End If
#If VBA7 Then
    Private Declare PtrSafe Function ShellExecuteW Lib "shell32" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpOperation As LongPtr, _
        ByVal lpFile As LongPtr, _
        ByVal lpParameters As LongPtr, _
        ByVal lpDirectory As LongPtr, _
        ByVal nShowCmd As Long) As Long
    'creates a complete directory of arbitrary depth:
    Private Declare PtrSafe Function MakeSureDirectoryPathExists Lib "imagehlp" (ByVal lpPath As LongPtr) As Long
#Else
    Private Declare Function ShellExecuteW Lib "shell32" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpOperation As LongPtr, _
        ByVal lpFile As LongPtr, _
        ByVal lpParameters As LongPtr, _
        ByVal lpDirectory As LongPtr, _
        ByVal nShowCmd As Long) As Long
    'erzeugt eine komplette Verzeichnisstruktur beliebiger Tiefe:
    Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp" (ByVal lpPath As LongPtr) As Long
#End If
'Public Enum FileAccess
Public Enum FileMode
    fmBinary_Read
    fmBinary_Write
    fmBinary_ReadWrite
    fmText_Read
    fmText_Write
End Enum
Private m_PFN As String
Private Const InvalidChars As String = "\/:*?""<>|" 'bei Dateiname noch "/\:"
Private Const C_DirSepChar As Long = 92 ' String = "\" ' = Chr(92)
Private Const C_ExtSepChar As Long = 46 ' String = "." ' = Chr(46)
Private Const C_InvalidChars As String = "*?""<>|" 'bei Dateiname noch "/\:"
Private Const C_3Pt As String = "..."

Private m_IsValid As Boolean

Private m_FNr As Integer
Private m_fm  As FileMode

' v ############################## v '    Constructors      ' v ############################## v '
Private Sub Class_Initialize()
End Sub
Private Sub Class_Terminate()
    Me.CloseFile
End Sub
Friend Sub New_(ByVal aPathOrPFN As String, _
       Optional ByVal aFileName As String, _
       Optional ByVal aExt As String)
        
    If Len(aFileName) = 0 And Len(aExt) = 0 Then
        m_PFN = RemoveQuotes(Trim(aPathOrPFN))  'RemoveQuotes(aPathFileName)
        m_IsValid = MakeValid(m_PFN)
    Else
        If Len(aExt) = 0 Then
            If Len(aPathOrPFN) = 0 Then
                m_PFN = GetValidFileName(RemoveQuotes(aFileName))
            Else
                m_PFN = GetValidPath(aPathOrPFN) & _
                        GetValidFileName(Get_FileName(aFileName))
            End If
        Else
            If Len(aFileName) Then
                m_PFN = GetValidPath(aPathOrPFN) & _
                        GetValidFileName(Get_FileName(aFileName)) & _
                        GetValidExt(aExt)
            Else
                'nur aPath und ext
                'aPath beinhaltet auch den Dateinamen
                'blöd und was wenn nich?
                'dann müßte man noch überprüfen ob der Pfad so überhaupt existiert
                'Naja, da gibts halt einfach nix anderes, wenn der Pfad falsch ist,
                m_PFN = MakeValid(RemoveQuotes(aPathOrPFN)) & GetValidExt(aExt)
            End If
        End If
    End If
End Sub
Friend Sub NewC(other As PathFileName) 'Copy-ctor
    m_PFN = other.Value
End Sub
Friend Function Clone() As PathFileName
    Set Clone = New PathFileName: Clone.Value = m_PFN
End Function

'internal ctor
Private Function New_PFN(ByVal aPathFileName As String, _
                Optional ByVal aFileName As String, _
                Optional ByVal aExt As String) As PathFileName
    Set New_PFN = New PathFileName: New_PFN.New_ aPathFileName, aFileName, aExt
End Function

' ^ ############################## ^ '    Constructors      ' ^ ############################## ^ '

' v ############################## v '  Public Properties   ' v ############################## v '
Public Property Get Value() As String
    Value = m_PFN
End Property
Public Property Let Value(RHS As String)
    m_PFN = RHS
    m_IsValid = True 'you know what you do, otherwise use New_
End Property

Public Property Get Quoted() As String
    Quoted = """" & m_PFN & """"
End Property

'a PathFileName consists of:
'              Path              |              FileName
'      C:\Windows\System32\      |             user32.dll
'resp:
'Drive     |      Path only      |   FileName only  |  Extension
'  C:\     |  Windows\System32\  |      user32      |    .dll
'


'Drive e.g.: "C:\" or "\\Server\" (UNC-path), drive will be closed by "\"
' * a local drive contains ":"
' * a network-drive is the name of a network-PC/server e.g "\\Server\"
Public Property Get Drive() As String
    Drive = Get_Drive(m_PFN)
End Property
Public Property Let Drive(ByVal Value As String)
    Let_Drive m_PFN, Value
End Property

Public Property Get PathOnly() As String
    PathOnly = Get_PathOnly(m_PFN)
End Property
Public Property Let PathOnly(ByVal Value As String)
    Let_PathOnly m_PFN, Value
End Property

Public Property Get Path() As String
    Path = Get_Path(m_PFN)
End Property
Public Property Let Path(ByVal Value As String)
    Let_Path m_PFN, Value
End Property

Public Property Get FileNameOnly() As String
    FileNameOnly = Get_FileNameOnly(m_PFN)
End Property
Public Property Let FileNameOnly(ByVal Value As String)
    Let_FileNameOnly m_PFN, Value
End Property

Public Property Get Extension() As String
    Extension = Get_Extension(m_PFN)
End Property
Public Property Let Extension(ByVal Value As String)
    Let_Extension m_PFN, Value
End Property

Public Property Get FileName() As String
    FileName = Get_FileName(m_PFN)
End Property
Public Property Let FileName(ByVal Value As String)
    Let_FileName m_PFN, Value
End Property

Public Property Get PathCount() As Long
    'returns the number of "\" in PathOnly
    Dim c As Long, i As Long
    Dim p As String: p = Me.PathOnly
    If Len(p) = 0 Then Exit Property
    For i = 1 To Len(p)
        If Mid$(p, i, 1) = "\" Then c = c + 1
    Next
    PathCount = c
    'PathCount = UBound(Split(Me.PathOnly), "\")
End Property
Public Property Get PathI(ByVal Index As Long) As String
    Dim p() As String: p = Split(Me.PathOnly, "\")
    If Index <= UBound(p) Then
        PathI = p(Index)
    End If
End Property
Public Property Let PathI(ByVal Index As Long, RHS As String)
    Dim p() As String: p = Split(Me.PathOnly, "\")
    If Index <= UBound(p) Then
        p(Index) = RHS
    End If
    Me.PathOnly = Join$(p, "\")
End Property

Public Property Get TempPath() As String
Try: On Error GoTo Catch
    TempPath = Environ("TMP")
    If Len(TempPath) = 0 Then TempPath = Environ("TEMP")
    TempPath = TempPath & "\"
    Exit Property
Catch:
    ErrHandler "TempPath"
End Property

Public Property Get UserPath() As String
    UserPath = Environ("USERPROFILE") & "\"
End Property

Public Property Get Attributes() As VbFileAttribute
    'works for existing files/pathes only!
    'Extras->Optionen->Allgemein->Unterbrechen bei Fehlern->(#)Bei nicht verarbeiteten Fehlern
Try: On Error GoTo Catch
    Attributes = GetAttr(m_PFN)
Catch:
End Property
'Public Property Get VersionInfo() As FileVersionInfo
'    'einfach auskommentieren wenn Version nicht gebraucht wird (in kleineren Projekten)
'    Set VersionInfo = New FileVersionInfo
'    Set VersionInfo = VersionInfo.GetVersionInfo(Me.Name)
'End Property
' ^ ############################## ^ '  Public Properties   ' ^ ############################## ^ '

' v ####################### v ' Public ReadOnly Boolean Properties ' v ####################### v '
Public Property Get IsIdent(other As Object) As Boolean
    IsIdent = ObjPtr(Me) = ObjPtr(other)
End Property
Public Property Get IsEqual(other As PathFileName) As Boolean
    IsEqual = Me.Value = other.Value
End Property
Public Property Get IsPath() As Boolean
    If Me.PathExists Then
        'works for existing pathes only!
        IsPath = (Me.Attributes And vbDirectory) = vbDirectory
    Else
        IsPath = Len(Me.FileName) = 0
    End If
End Property
Public Property Get IsFile() As Boolean
    IsFile = Len(Me.FileName) > 0
End Property
Public Property Get IsUNC() As Boolean
    Dim d As String: d = Me.Drive
    If Len(d) = 0 Then Exit Property
    IsUNC = Left$(d, 2) = "\\"
End Property
Public Property Get IsValid() As Boolean
    IsValid = m_IsValid
End Property
' ^ ####################### ^ ' Public ReadOnly Boolean Properties ' ^ ####################### ^ '


' v ############################## v '   Public Functions   ' v ############################## v '

#If VBA7 Then
    Public Function Start(Optional ByVal Cmd As String = "", Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus) As Long
Try:     On Error GoTo Catch
        'startet ohne zu warten = asynchron
        Dim ret As Long
        If Me.FileName = "" Then
            Dim p As PathFileName: Set p = Me
            Do While Not p Is Nothing
                If p.PathExists Then Exit Do
                'If Not p.PathExists Then
                Set p = p.PathUp
            Loop
            If Not p Is Nothing Then
                ret = Shell("Explorer.exe " & p.Path, vbNormalFocus)
            End If
            Exit Function
        Else
            Dim e As String: e = UCase(Me.Extension)
            If e = "EXE" Or e = "BAT" Then
                'ein ausführbares programm
                Cmd = IIf(Len(Cmd) = 0, m_PFN, m_PFN & " " & Cmd)
                ret = Shell(Cmd, WindowStyle) ' liefert die TaskID
            Else
                'ein Dokument
                Dim hwnd As LongPtr
                'Debug.Print m_PFN
                ret = ShellExecute(hwnd, "Open", Me.Quoted, Cmd, Me.Path, WindowStyle)
            End If
        End If
        Start = ret
        Exit Function
Catch:
        ErrHandler "Start"
    End Function
    
    Public Function StartWait(Optional ByVal Cmd As String = "", Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, Optional ByVal WorkDir As String) As Long
Try:     On Error GoTo Catch
        'startet und wartet = synchron
        'geht jetzt nur bei exe (oder bat?)
        Dim e As String: e = UCase(Me.Extension)
        If Not (e = "EXE" Or e = "BAT") Then
            MsgBox "Nur ausführbare Programme (exe) können gestartet werden, Dateiendung: " & Me.Extension
            Exit Function
        End If
        Dim ret As LongPtr
        Cmd = IIf(Len(Cmd) = 0, m_PFN, m_PFN & " " & Cmd)
        ret = MShell.ShellWait(Cmd, WindowStyle, WorkDir)
        'StartWait = ret
        Exit Function
Catch:
        ErrHandler "StartWait"
    End Function
    
#Else 'VBA6, VBC

    Public Function Start(Optional ByVal Cmd As String = "", Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus) As Long
Try:     On Error GoTo Catch
        'startet ohne zu warten = asynchron
        Dim ret As Long
        If Me.FileName = "" Then
            Dim p As PathFileName: Set p = Me
            Do While Not p Is Nothing
                If p.PathExists Then Exit Do
                'If Not p.PathExists Then
                Set p = p.PathUp
            Loop
            If Not p Is Nothing Then
                ret = Shell("Explorer.exe " & p.Path, vbNormalFocus)
            End If
            Exit Function
        Else
            Dim e As String: e = UCase(Me.Extension)
            If e = "EXE" Or e = "BAT" Then
                'ein ausführbares programm
                Cmd = IIf(Len(Cmd) = 0, m_PFN, m_PFN & " " & Cmd)
                ret = Shell(Cmd, WindowStyle) ' liefert die TaskID
            Else
                'ein Dokument
                Dim hwnd As Long
                'Debug.Print m_PFN
                ret = ShellExecuteW(hwnd, StrPtr("Open"), StrPtr(Me.Quoted), StrPtr(Cmd), StrPtr(Me.Path), WindowStyle)
            End If
        End If
        Start = ret
        Exit Function
Catch:
        ErrHandler "Start"
    End Function
    
    Public Function StartWait(Optional ByVal Cmd As String = "", Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, Optional ByVal WorkDir As String) As Long
Try:     On Error GoTo Catch
        'startet und wartet = synchron
        'geht jetzt nur bei exe (oder bat?)
        Dim e As String: e = UCase(Me.Extension)
        If Not (e = ".EXE" Or e = ".BAT" Or e = ".COM") Then
            MsgBox "Start executable programs only (like .exe, .bat or .com), your file extension: " & Me.Extension
            Exit Function
        End If
        Dim ret As Long
        Cmd = IIf(Len(Cmd) = 0, m_PFN, m_PFN & " " & Cmd)
        ret = MShell.ShellWait(Cmd, WindowStyle, WorkDir)
        'StartWait = ret
        Exit Function
Catch:
        ErrHandler "StartWait"
    End Function

#End If

Public Function RemoveQuotes(Optional ByVal aStrPFN As String) As String
    Dim q1 As Long: q1 = IIf(Asc(Left$(aStrPFN, 1)) = 34, 2, 1)
    Dim q2 As Long: q2 = IIf(Asc(Right$(aStrPFN, 1)) = 34, Len(aStrPFN) - 1, Len(aStrPFN))
    RemoveQuotes = Mid$(aStrPFN, q1, q2 - q1 + 1)
End Function

Public Function Shorted(Optional ByVal MaxLen As Long) As String
    'returns the filenam shorted     'gibt einen gekürzten String zurück
    'for the purpose of showing a filename in program file-menu if the filename is to long for the menu.
    If MaxLen = 0 Then MaxLen = 100
    'wenn der Dateiname länger als 100 ist, dann nur den Dateiname anzeigen
    Shorted = Me.FileNameOnly
    If Len(Shorted) < MaxLen Then
        'Wenn kleiner dann noch das Laufwerk dazunehmen:
        Shorted = Me.Drive & "..." & Me.FileNameOnly
        If Len(Shorted) < MaxLen Then
            'wenn kleiner dann noch die Dateierweiterung dazunehmen
            Shorted = Me.Drive & "..." & Me.FileName
            If Len(Shorted) < MaxLen Then
                'wenn kleiner dann den Pfad dazunehmen
                'für den verbleibenden Rest den Pfad in zwei Teile teilen
                Dim d As Long: d = MaxLen - Len(Shorted)
                If d > Len(PathOnly) Then
                'wenn Pfad kleiner als der Rest dann zeige ganzen Dateinamen
                    Shorted = Me.Value
                Else
                    Dim half As Long: half = d \ 2 - 2 '-2 wegen den Pünktchen
                    If half > 0 Then
                        Dim o As String: o = Me.PathOnly
                        Shorted = Left$(o, half) & "..." & Right$(o, half)
                    Else 'dann nur die Pünktchen ausgeben
                        Shorted = "..."
                    End If
                    Shorted = Me.Drive & Shorted & Me.FileName
                End If
            End If
        End If
    End If
End Function

Public Function pStr(Optional ByVal BufLen As Long = 1024) As Long
    'for working with some win-api functions
    'returns the pointer to the filename-string.
    '
    'übergibt den Zeiger auf den String falls noch kein Wert vorhanden, wird ein
    'Puffer angelegt, und der Zeiger darauf übergeben.
    If Len(m_PFN) = 0 Then
        m_PFN = String$(BufLen, vbNullChar)
    End If
    pStr = StrPtr(m_PFN)
End Function
'
'Private Function GetValidExtension(Value As String) As String
'    GetValidExtension = Replace(Trim$(Value), ":\/", "_")
'End Function

Public Function PathCreate(Optional ByVal sNewDir As String) As Boolean
TryE: On Error GoTo CatchE
    'ein neues Verzeichnis mit Unterverzeichnissen anlegen
    If Len(sNewDir) = 0 Then sNewDir = Me.Path
    Call MakeValid(sNewDir)
    Dim NewDir As PathFileName: Set NewDir = New_PFN(sNewDir)
    'zuerst in das jeweilige Laufwerk wechseln
    Dim sd As String: sd = NewDir.Drive()
    Call ChDir(sd)
    Dim i As Long
    For i = 0 To NewDir.PathCount - 1
        sd = NewDir.PathI(i)
        'das Verzeichnis anlegen
        Call MkDir(sd)
        'gleich ins verzeichnis wechseln, damit das nächste
        'Unterverzeichnis angelegt werden kann
        Call ChDir(sd)
    Next
    PathCreate = True
    Exit Function
CatchE:
    ErrHandler "PathCreate", "Could not create path: " & vbCrLf & sNewDir
End Function
'
'Public Function PathCreate() As Boolean
'Try: On Error GoTo Catch
'    'erzeugt den Pfad und gibt zurück obs geklappt hat, bzw ob der Pfad dann existiert
'    'wie reagiert die Funktion wenn
'    ' * Pfad bereits vorhanden
'    ' * Pfad mit mehreren Unterverzeichnissen, werden alle Unterverzeichnisse automatisch gleichzeitig mit angelegt?
'    ' * oder wird nur immer ein Unterverzeichnis angelegt?
'    ' * was ist mit "MakeSureDirectoryPathExists"?
'    '
'    MkDir Me.Path
'    PathCreate = Me.PathExists
'    Exit Function
'Catch:
'    ErrHandler "PathCreate"
'End Function

Public Function PathDelete(Optional ByVal sDelDir As String) As Boolean
Try: On Error GoTo Catch
    If Len(sDelDir) = 0 Then
        sDelDir = Me.Path
        If Me.Exists() Then Exit Function
    End If
    'Call RemoveDirectoryA(deldir)
    'nur ein komplett leeres Directory
    Call RmDir(sDelDir)
    PathDelete = True
    Exit Function
Catch:
    ErrHandler "PathDelete", sDelDir
End Function

Public Function PathUp() As PathFileName
    Dim p As String: p = m_PFN
    If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
    Dim pos As Long: pos = InStrRev(p, "\")
    If pos > 0 Then
        Set PathUp = New_PFN(Left(p, pos))
    End If
End Function

Public Function CheckExtensions(ParamArray exts()) As String
    'checks whether the file extension is contained in the array exts()
    'überprüft ob die extension in dem angegebenen Array exts() enthalten ist
    Dim e As String: e = Me.Extension
    Dim i As Integer
    For i = 0 To UBound(exts)
        'wenn mindestens eine zutrifft dann OK, es wird die Extension zurückgegeben, andernfalls ""
        If e = exts(i) Then
            CheckExtensions = e
            Exit Function
        End If
    Next
End Function

Public Function ChangeExt(newExtension As String) As PathFileName
    Set ChangeExt = Me.Clone
    ChangeExt.Extension = newExtension
'    If Len(newExtension) > 0 Then
'        If Left(newExtension, 1) <> "." Then newExtension = "." & newExtension
'    End If
'    Set ChangeExt = New PFN: ChangeExt.New_ Me.Path & Me.FileNameOnly & newExtension
End Function
Public Function AddToFileNameR(ByVal addstr As String) As PathFileName
    Dim e As String: e = Me.Extension 'Änderung 08.01.2016
    Set AddToFileNameR = New_PFN(Me.Path & Me.FileNameOnly & addstr & IIf(Len(e), "." & e, ""))
End Function
Public Function AddToFileNameL(ByVal addstr As String) As PathFileName
    Dim e As String: e = Me.Extension 'Änderung 08.01.2016
    Set AddToFileNameL = New_PFN(Me.Path & addstr & Me.FileNameOnly & IIf(Len(e), "." & e, ""))
End Function

Public Function Exists() As Boolean
    'checks whether file or path exists, works on both(!)
    'Überprüft ob die Datei oder der Pfad existiert
    On Error Resume Next
    Dim f As String: f = Me.FileName
    If Len(f) Then
    'If Me.IsFile Then
        Exists = Not CBool(GetAttr(m_PFN) And (vbDirectory Or vbVolume))
    Else
        Dim p As String: p = Me.Path
        Exists = Exists = CBool(GetAttr(p) And vbDirectory)
    End If
    On Error GoTo 0
End Function

Public Function PathExists() As Boolean
    'checks only whether the path exists
    'Überprüft nur ob das Verzeichnis existiert
    On Error Resume Next
    Dim p As String: p = Me.Path
    'Extras->Optionen->Allgemein->Unterbrechen bei Fehlern->(*) Bei nicht verarbeiteten Fehlern
    'Laufzeitfehler 76 Pfad nicht gefunden
    'Runtime-Error  76 Path not found
    'just go ahead
    If Not IsInIDE Then
        PathExists = CBool(GetAttr(p) And vbDirectory)
    End If
    On Error GoTo 0
End Function


'  Set fso = CreateObject("Scripting.FileSystemObject")
'   gefunden = False
'   For Each l In fso.Drives
'     If l.DriveType = 3 Then
'       If l.ShareName = "\\Netzwerkpfad\jobs" Then
'         gefunden = True: lw = l.DriveLetter & ":\"
'       End If
'     End If
'   Next
'   If gefunden = False Then
'     On Error Resume Next
'     Set objNetzwerk = CreateObject("WScript.Network")
'     objNetzwerk.MapNetworkDrive "B:", "\\Netzwerkpfad\jobs"
'     If Err.Number <> 0 Then
'       MsgBox Err.Number & vbTab & Err.Description
'     Else
'       MsgBox "Verbindung mit \\Netzwerkpfad\jobs auf B:\ hergestellt"
'     End If
'   Else
'     MsgBox lw & " mit \\Netzwerkpfad\jobs verbunden"
'   End If

'OM: ÄNDERUNG 08.06.2016 JETZT WIRD AUCH KOPIERT WENN NEWPATH NUR EIN PFAD IST, ES WIRD DER DATEINAME ANGEHÄNGT
'Halt, ist das evtl fehlerhaft?
'SHIT, wie ist CopyTo definiert? soll der gleiche Dateiname in eine verzeichnis kopiert werden, oder kann ein Dateiname angegeben werden

'File operations: Copy, Move, Delete, Rename, MakeDir, RemoveDir


'OK Frage:
'Sollen bestimmte Funktionen ein PFN-Objekt zurückgeben, oder lieber nur einfache Typen wie Boolean oder String?
'gibt man ein PFN-Objekt zurück so kann man mehrere Operationen in einer Zeile erledigen.
Public Function CopyTo(newPathFileName As PathFileName, Optional bOverwrite As Boolean = True) As PathFileName
'per default wird immer die Datei überschrieben => "bOverwrite As Boolean = True"
'nur wenn bOverwrite extra mit False angegeben wird, wird nicht überschrieben
Try: On Error GoTo Catch
    Dim newPFN As PathFileName
    If newPathFileName.IsPath Then
        Set newPFN = New_PFN(newPathFileName.Path & Me.FileName)
    Else
        Set newPFN = newPathFileName
    End If
    'FileCopy strOldFile, strNewFile
    If Not bOverwrite Then
        'Debug.Print newPathFileName.Value
        If newPFN.Exists Then Exit Function
    End If
    'OK keep filename in case newPathFileName is not given
    If Not newPFN.PathExists Then
        MsgBox "der Pfad existiert nicht: " & vbCrLf & newPFN.Path
        Exit Function
    End If
    FileCopy m_PFN, newPFN.Value
    Set CopyTo = newPFN
    Exit Function
Catch:
    If Err.Number = 70 Or Err.Number = 53 Then
        'File access denied? give it another try one or more, MsgBox retry
        Set CopyTo = CopyRetryLoop(newPFN, True)
    Else
        ErrHandler "CopyTo", vbCrLf & m_PFN & vbCrLf & vbCrLf & newPFN.Value
    End If
End Function

Private Function CopyRetryLoop(newPathFileName As PathFileName, Optional bOverwrite As Boolean = True) As PathFileName
Try:  On Error GoTo Catch
    FileCopy m_PFN, newPathFileName.Value
    Set CopyRetryLoop = newPathFileName
    Exit Function
Catch:
    Dim rv As VbMsgBoxResult: rv = MessErrorRetry(Me, "CopyTo", m_PFN & vbCrLf & newPathFileName.Value)
    If rv = vbRetry Then
        GoTo Try 'yes we use goto here, since the user has the choice to stop it anytime
    Else
        Set CopyRetryLoop = Nothing
        Exit Function
    End If
End Function

Public Function MoveTo(pfndst As PathFileName) As PathFileName
    'move = nach Dst kopieren und Src löschen
    'Move = verschieben, im Quellordner ist die Datei weg
    'Copy = Kopieren,    im Quellordner ist die Datei noch da
Try: On Error GoTo Catch
    Set MoveTo = Me.CopyTo(pfndst)
    If MoveTo Is Nothing Then Exit Function
    Me.Delete
    Exit Function
Catch:
    ErrHandler "MoveTo", m_PFN & vbCrLf & pfndst.Value
End Function

Public Function Delete(Optional ByVal aPFN As String = vbNullString) As Boolean
'löscht nur was es auch gibt
Try: On Error GoTo Catch
    If aPFN = vbNullString Then aPFN = m_PFN
    Kill aPFN
    Delete = True
    Exit Function
Catch:
    ErrHandler "Delete", aPFN
End Function

'ist BolRenameMe = False, dann wird dieses Objekt nicht verändert,
'sondern nur die Datei auf der Festplatte umbenannt
'dieses Objekt existiert dann nicht mehr als Datei auf der Festplatte
'(FileExists liefert dann false)
Public Function Rename(ByVal ToPFN As String, _
                       Optional ByVal BolRenameMe As Boolean) As Boolean
TryE: On Error GoTo CatchE
    Dim RenFile As PathFileName: Set RenFile = New_PFN(ToPFN)
    Dim Cmd     As String
    If Not RenFile.Exists Then
        'Wird nicht renamed wenn es die Datei schon gibt
        Cmd = Chr(34) & Me.Value & Chr(34) & " " & Chr(34) & RenFile.Value & Chr(34)
        Call Shell("cmd /c ren " & Cmd, 0)
        Rename = True
        If BolRenameMe Then
            'soll nur die Datei auf der Festplatte umbenannt werden
            'oder auch dieser Dateiname selber verändert werden
            Me.Value = ToPFN
        End If
    End If
    Exit Function
CatchE:
    ErrHandler "Rename", ToPFN
End Function

'Functions for Writing to and Reading from files
Public Function OpenFile(fm As FileMode) As Integer
    m_fm = fm
Try: On Error GoTo Catch
    If m_FNr <> 0 Then OpenFile = m_FNr: Exit Function
    m_FNr = FreeFile
    Select Case m_fm
    Case fmBinary_Read
        
        Open m_PFN For Binary Access Read As m_FNr
        
    Case fmBinary_Write
        
        Open m_PFN For Binary Access Write As m_FNr
        
    Case fmBinary_ReadWrite
        
        Open m_PFN For Binary As m_FNr
        
    Case fmText_Read
        
        Open m_PFN For Input As m_FNr
        
    Case fmText_Write
        
        Open m_PFN For Output As m_FNr
        
    End Select
    OpenFile = m_FNr
    Exit Function
Catch:
    ErrHandler "OpenFile", m_PFN
End Function

Public Sub WriteStr(aStr As String)
Try: On Error GoTo Catch
    If m_FNr = 0 Then OpenFile fmBinary_Write
    Select Case m_fm
    Case fmBinary_Write
        Put #m_FNr, , aStr
    Case fmText_Write
        Write #m_FNr, aStr
    End Select
    Exit Sub
Catch:
    ErrHandler "WriteStr", m_PFN
End Sub

Public Function ReadStr() As String
Try: On Error GoTo Catch
    If m_FNr = 0 Then OpenFile fmBinary_Read
    Dim aStr As String
    Select Case m_fm
    Case FileMode.fmBinary_Read
        aStr = Space(LOF(m_FNr))
        Get #m_FNr, , aStr
    Case FileMode.fmText_Read
        Input #m_FNr, aStr
    End Select
    ReadStr = aStr
    Exit Function
Catch:
    ErrHandler "ReadStr", m_PFN
End Function

Public Sub CloseFile()
    If m_FNr = 0 Then Exit Sub
Try: On Error GoTo Catch
    Close #m_FNr
    m_FNr = 0
    Exit Sub
Catch:
    ErrHandler "CloseFile", m_PFN
End Sub
' ^ ############################## ^ '   Public Functions   ' ^ ############################## ^ '

' v ############################## v '   Private Parsers    ' v ############################## v '
Private Function Get_Drive(PFN_in As String) As String
    'can be either like "C:\" or like "\\Server\" (UNC);
    'a local drive contains ":", a network-drive contains "\\" at the beginning;
    'and yes, a linked network drive will be treated here as a local drive
    'it will always be closed by "\"
    Dim s As String
    Dim pos As Long: pos = InStr(1, PFN_in, ":")
    If 0 < pos And pos < 3 Then
        s = Left$(PFN_in, pos) & "\"
    ElseIf Left$(PFN_in, 2) = "\\" Then
        pos = InStr(3, PFN_in, "\")
        If pos = 0 Then
            If Right$(PFN_in, 1) <> "\" Then s = PFN_in & "\"
        ElseIf pos > 2 Then
            s = Left$(PFN_in, pos)
        End If
    End If
    Get_Drive = s
End Function
Private Sub Let_Drive(ByRef PFN_out As String, ByVal Value As String)
    'OM TODO: make valid Drive from given Value
    Value = Trim$(Value)
    PFN_out = Value & Get_PathOnly(PFN_out) & Get_FileName(PFN_out)
End Sub

Private Function Get_PathOnly(PFN_in As String) As String
    Dim d As String: d = Get_Drive(PFN_in)
    Dim p As String: p = Get_Path(PFN_in)
    If Len(p) Then Get_PathOnly = Mid$(p, Len(d) + 1)
End Function
Private Sub Let_PathOnly(ByRef PFN_out As String, ByVal Value As String)
    'OM TODO: make valid PathOnly from given Value
    Value = Trim$(Value)
    PFN_out = Get_Drive(PFN_out) & Value & Get_FileName(PFN_out)
End Sub

Private Function Get_Path(PFN_in As String) As String
'Try: On Error GoTo Catch
    Dim s As String
    Dim d As String: d = Get_Drive(PFN_in)
    If Len(PFN_in) <= Len(d) Then
        s = d
    Else
        'rbs_pos: position of rightmost backslash "\"
        Dim rbs_pos As Long: rbs_pos = InStrRev(PFN_in, "\")
        If rbs_pos = 0 Then Exit Function
        s = Left$(PFN_in, rbs_pos) 'mit "\"
    End If
    Get_Path = s
'    Exit Function
'Catch:
'    If Err Then MessError Me, "Path"
End Function
Private Sub Let_Path(ByRef PFN_out As String, ByVal Value As String)
    'OM TODO: make valid Drive and Path from given Value
    'does Value contain a Drive Letter?
    'if not add the existing one
    Value = Trim$(Value)
    PFN_out = Value & FileName
End Sub

Private Function Get_FileNameOnly(PFN_in As String) As String
    'gibt den Dateinamen ohne Drive, ohne Path und ohne Extension zurück
'Try: On Error GoTo Catch
    Dim f As String: f = Get_FileName(PFN_in)
    Dim e As String: e = Get_Extension(PFN_in)
    If Len(f) Then
        If Len(e) Then
            f = Left(f, Len(f) - Len(e))
        End If
    End If
    Get_FileNameOnly = f
'Catch:
'    If Err Then MessError Me, "FileNameOnly"
End Function
Private Sub Let_FileNameOnly(ByRef PFN_out As String, ByVal Value As String)
    'ohne Pfad und Extension
    Value = Trim$(Value)
    m_PFN = Me.Path & Value & Me.Extension
End Sub

Private Function Get_Extension(PFN_in As String) As String
    Dim rbs_pos As Long: rbs_pos = InStrRev(PFN_in, "\")  'position of rightmost backslash "\"
    Dim rpt_pos As Long: rpt_pos = InStrRev(PFN_in, ".")  'position of rightmost point "."
    If rpt_pos = 0 Then Exit Function
    If rpt_pos < rbs_pos Then Exit Function
    Get_Extension = Mid$(PFN_in, rpt_pos)
End Function
Private Sub Let_Extension(ByRef PFN_out As String, ByVal Value As String)
    Value = Trim$(Value)
    If Left(Value, 1) <> "." Then Value = "." & Value
    m_PFN = Me.Path & Me.FileNameOnly & Value
End Sub

Private Function Get_FileName(PFN_in As String) As String
    Dim p As String: p = Get_Path(PFN_in)
    Get_FileName = Mid$(PFN_in, Len(p) + 1)
End Function
Private Sub Let_FileName(ByRef PFN_out As String, ByVal Value As String)
    PFN_out = Get_Path(PFN_out) & Value
End Sub

Private Function Get_DirPart(ByVal n As Long, _
                             ByRef pos1 As Long, _
                             ByRef pos2 As Long, _
                             ByRef sDir As String) As Boolean
'gibt einen Teil des Directories  zurück
'OK man könnte das Ganze auch mit einem Array machen.
'Aber da man normalerweise nur selten einzelne Teile
'des Pfades ändert, lohnt sich der Verwaltungsaufwand
'nicht.
    Dim op As String: op = GetValidPath(Me.PathOnly)
    Dim i  As Long
    Dim P1 As Long, P2 As Long
    For i = 0 To n - 1
        P1 = P2 + 1
        P2 = InStr(P1, op, Chr$(C_DirSepChar))
    Next
    If P2 > P1 Then
        sDir = Mid$(op, P1, P2 - P1)
        pos1 = P1
        pos2 = P2
        Get_DirPart = True
    End If
End Function
' ^ ############################## ^ '   Private Parsers    ' ^ ############################## ^ '

' v ############################## v '  Private Validators  ' v ############################## v '

'was ist der Punkt der Validierer, wie arbeiten die mit sich und mit den Props zusammen?

'Ein Dateiname darf keines der folgenden Zeichen enthalten: \/:*?"<>|
'Valid "/" durch "\" und "*?"<>|" durch "_"

Private Function GetValidDrive(aDrive As String) As String
'falls nicht ganz rechts vorhanden,
'stattet eine Laufwerksangabe mit einem Pfad-Trennzeichen ("\") aus.
'überprüft ob es ein lokales Laufwerk (":") oder ein Nwetzwerk-laufwerk ("\\") ist
'entfernt alle Vorkommen von Whitespaces " ", Tab etc.
    If LenB(aDrive) > 0 Then
        If Asc(Mid$(aDrive, Len(aDrive) - 1)) <> C_DirSepChar Then
            GetValidDrive = aDrive & Chr$(C_DirSepChar)
        Else
            GetValidDrive = aDrive
        End If
    End If
End Function
Private Function GetValidPath(ByVal aPath As String) As String
'falls nicht ganz rechts vorhanden,
'stattet einen Pfad mit einem Pfad-Trennzeichen ("\") aus.
    If LenB(aPath) > 0 Then
        If Asc(Mid$(aPath, Len(aPath))) <> C_DirSepChar Then
            GetValidPath = aPath & Chr$(C_DirSepChar)
        Else
            GetValidPath = aPath
        End If
    End If
End Function
Private Function IsValidFileName(aFileName As String) As Boolean
    Dim i As Long
    Dim X As String: X = "\/:*?""<>|"
    If Len(aFileName) Then
        For i = 1 To Len(X)
            IsValidFileName = (InStr(1, aFileName, Mid$(X, i, 1)) <> 0)
            If Not IsValidFileName Then Exit Function
        Next
    End If
End Function
Private Function GetValidFileName(aFileName As String) As String
    Dim i As Long
    Dim c As String, X As String: X = "\/:" & C_InvalidChars '*?""<>|"
    If Len(aFileName) Then
        'GetValidFileName = aFileName
        GetValidFileName = RemoveChars(aFileName, X)
        'For i = 1 To Len(x)
        '    c = Mid$(x, i, 1)
        '    GetValidFileName = Replace$(GetValidFileName, c, vbNullString)
        'Next
    End If
End Function
Private Function GetValidExt(aExt As String) As String
'falls nicht ganz links vorhanden,
'stattet eine  Erweiterung mit einem Extensions-Trennzeichen (".") aus.
    If Len(aExt) > 0 Then
        Dim c As String: c = Chr$(C_ExtSepChar)
        If InStr(1, aExt, c) = 0 Then
            GetValidExt = c & aExt
        Else
            GetValidExt = aExt
        End If
    End If
End Function
Private Function GetDividedPath(MaxLen As Long) As String
    Dim half As Long: half = MaxLen \ 2 - 2 '-2 wegen den Pünktchen
    If half > 0 Then
        Dim o As String: o = PathOnly
        GetDividedPath = Left$(o, half) & C_3Pt & Right$(o, half)
    Else 'dann nur die Pünktchen ausgeben
        GetDividedPath = C_3Pt
    End If
End Function
Private Function RemoveChars(ByVal s As String, CharsToRemove As String) As String
    Dim c As String
    Dim i As Long
    RemoveChars = s
    For i = 1 To Len(CharsToRemove)
        c = Mid$(CharsToRemove, i, 1)
        If InStr(1, s, c) Then
            RemoveChars = Replace$(RemoveChars, c, vbNullString)
        End If
    Next
End Function
'https://de.wikipedia.org/wiki/Uniform_Naming_Convention
'

Private Function MakeValid(ByRef aStrPath As String) As Boolean
    'Nullen von hinten entfernen
Try: On Error GoTo Catch
    aStrPath = Trim$(aStrPath)
    Dim p As Long: p = InStr(1, aStrPath, vbNullChar)
    If p <> 0 Then aStrPath = Left$(aStrPath, p - 1)
    'und alle Zeichen entfernen die nicht hinein dürfen
    Dim i As Long
    Dim c As String, X As String: X = C_InvalidChars '"*?""<>|"
    Select Case C_DirSepChar
    Case 92: X = X & "/"
    Case 47: X = X & "\"
    End Select
    aStrPath = RemoveChars(aStrPath, X)
    aStrPath = RecursiveReplaceSL(aStrPath, "\\", "\", 2)
'    'und alle doppelten Backslashes entfernen,
'    'die nicht am Anfang sind, denn ansonsten ist es ein Serverlaufwerk
    'braucht man das wirklich? nein vielleicht gibt es da Fehler irgendwo
    c = Chr$(C_DirSepChar)
    X = c & c
    Dim posDblBackSlash As Long
    posDblBackSlash = InStr(1, aStrPath, X)
    If posDblBackSlash > 1 Then
        aStrPath = Replace(aStrPath, X, c)
    End If
    'und doppelpunkte ":" entfernen ausser in der Laufwerksangabe
    Dim sd As String: sd = Get_Drive(aStrPath)
    Dim ld As Long:   ld = Len(sd)
    c = Chr$(58) '":"
    If InStr(ld + 1, aStrPath, c) Then
       'grummel grummel, replace gibt nur einen Teil zurück
        aStrPath = sd & Replace$(aStrPath, c, vbNullString, ld + 1)
    End If
    MakeValid = True
    Exit Function
Catch:
    '
End Function
' ^ ############################## ^ '  Private Validators  ' ^ ############################## ^ '

' v ############################## v '   Local ErrHandler   ' v ############################## v '
Private Function ErrHandler(ByVal PrcName As String, _
                            Optional AddInfo As String, _
                            Optional BolLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKOnly Or vbCritical, _
                            Optional bRetry As Boolean) As VbMsgBoxResult
    'ErrHandler = GlobalErrHandler(Me, PrcName, AddInfo, BolLoud, bErrLog, vbDecor)
    If bRetry Then
        ErrHandler = MessErrorRetry(TypeName(Me), PrcName, AddInfo, bErrLog)
    Else
        ErrHandler = MessError(TypeName(Me), PrcName, AddInfo, BolLoud, bErrLog, vbDecor)
    End If
End Function


'Die Klasse PFN alias PathFileName
'dient
'Dokumentation, was soll zurückgeliefert werden
'grundsätzlich gilt die Windows-Vorgabe:
'Ein Dateiname darf keines der folgenden Zeichen enthalten: \/:*?"<>|
'Valid "/" durch "\" und "*?"<>|" durch "_"
'es gibt grundsätzlich 2 unterschiedliche Arten von Pfaden
' * klassischer lokaler Pfad:
'         z.B.        : "C:\Hauptverzeichnis\Unterverzeichnis\Datei.txt"
'         Drive       : "C:\" muss ":" enthalten und muss rechts ein "\" enthalten
'         Path        : "C:\Hauptverzeichnis\Unterverzeichnis\" muss rechts ein "\" enthalten
'         PathCount = 2; PathI(0) = "Hauptverzeichnis\"; PathI(1) = "Unterverzeichnis\";
'         FileName    :                                      "Datei.txt"
'         Extension   :                                           ".txt" muss links ein "." enthalten
'         PathOnly    :    "Hauptverzeichnis\Unterverzeichnis\"
'         FileNameOnly:                                      "Datei"
'
' * UNC-Pfad auf einem Rechner im Netzwerk
'         z.B.        : "\\Server\Hauptverzeichnis\Unterverzeichnis\Datei.txt"
'         Drive       : "\\Server\" muss "\\" enthalten
'         Path        : "\\Server\Hauptverzeichnis\Unterverzeichnis\"
'         PathCount = 2; PathI(0) = "Hauptverzeichnis\"; PathI(1) = "Unterverzeichnis\";
'         FileName    :                                            "Datei.txt"
'         Extension   :                                                 ".txt"
'         PathOnly    :          "Hauptverzeichnis\Unterverzeichnis\"
'         FileNameOnly:                                            "Datei"
'
'wenn Teile fehlen
'       * z.B.        : "Datei"
'         Drive       : ""
'         Path        : ""
'         FileName    : "Datei"
'         Extension   : ""
'         PathOnly    : ""
'         FileNameOnly: "Datei"
'
'       * z.B.        : "Datei.txt"
'         Drive       : ""
'         Path        : ""
'         FileName    : "Datei.txt"
'         Extension   : ".txt"
'         PathOnly    : ""
'         FileNameOnly: "Datei"
'
'       * z.B.        : "Hauptverzeichnis\Unterverzeichnis\"
'         Drive       : ""
'         Path        : "Hauptverzeichnis\Unterverzeichnis\"
'         PathCount = 2; PathI(0) = "Hauptverzeichnis\"; PathI(1) = "Unterverzeichnis\";
'         FileName    : ""
'         Extension   : ""
'         PathOnly    : "Hauptverzeichnis\Unterverzeichnis\"
'         FileNameOnly: ""
'
'       * z.B.        : "C:" oder "C:\"
'         Drive       : "C:\"
'         Path        : ""
'         FileName    : ""
'         Extension   : ""
'         PathOnly    : ""
'         FileNameOnly: ""
'
'       * z.B.        : "\\Server" oder "\\Server\"
'         Drive       : "\\Server\"
'         Path        : ""
'         FileName    : ""
'         Extension   : ""
'         PathOnly    : ""
'         FileNameOnly: ""
'
'       * z.B.        : ".txt"
'         Drive       : ""
'         Path        : ""
'         FileName    : ".txt"
'         Extension   : ".txt"
'         PathOnly    : ""
'         FileNameOnly: ""


