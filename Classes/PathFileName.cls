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
'For this class to compile you also need:
'* MShell.bas  from the same repo                     in folder \Modules\
'* MString.bas from the repo Sys_Strings              in folder \Modules\
'* MPtr.bas    from the repo Ptr_Pointers             in folder \Modules\
'* MErr.bas    from the repo Err_CorrectErrorHandling in folder \Modules\
'
'a PathFileName consists of the following parts:
'Drive     |      Only path      |   Only filename  |  extension
'  C:\     |  Windows\System32\  |      user32      |    .dll
'
'resp:
'      Path (=Drive+OnlyPath)    |  Filename
'  C:\Windows\System32\          |  user32.dll

Option Explicit
'OM 06.11.2015 13:45; Zeilen:  336
'OM 2022-01-03 15:46; Zeilen: 1386
'OM 2022-11-01 16:29; Zeilen: 1773
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
    Private Declare PtrSafe Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As LongPtr) As Long
    Private Declare PtrSafe Function SetFileAttributesW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal Value As Long) As Long
#Else
    'activate the Enum LongPtr if you do not want to declare it in another module, maybe because this is the only class you use
    'Private Enum LongPtr
    '    [_]
    'End Enum
    Private Declare Function ShellExecuteW Lib "shell32" ( _
        ByVal hwnd As LongPtr, _
        ByVal lpOperation As LongPtr, _
        ByVal lpFile As LongPtr, _
        ByVal lpParameters As LongPtr, _
        ByVal lpDirectory As LongPtr, _
        ByVal nShowCmd As Long) As Long
    'erzeugt eine komplette Verzeichnisstruktur beliebiger Tiefe:
    Private Declare Function MakeSureDirectoryPathExists Lib "imagehlp" (ByVal lpPath As LongPtr) As Long
    Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As LongPtr) As Long
    Private Declare Function SetFileAttributesW Lib "kernel32" (ByVal lpFileName As LongPtr, ByVal Value As Long) As Long
#End If

Enum EFileAttribute
    None = VbFileAttribute.vbNormal         '         0
    ReadOnly = VbFileAttribute.vbReadOnly   '         1 Die Datei ist schreibgeschützt. ReadOnly wird unter Windows, Linux und macOS unterstützt. Unter Linux und macOS ist das Ändern des ReadOnly-Flags ein Vorgang mit Berechtigungen.
    Hidden = VbFileAttribute.vbHidden       '         2 Die Datei ist versteckt und daher nicht in einer normalen VerzeichnisListe enthalten. Hidden wird unter Windows, Linux und macOS unterstützt.
    System = VbFileAttribute.vbSystem       '         4 Die Datei ist eine Systemdatei. Die Datei ist also Teil des Betriebssystems oder wird ausschließlich durch das Betriebssystem verwendet.
    Volume = VbFileAttribute.vbVolume       '         8  '???
    Directory = VbFileAttribute.vbDirectory '&H10 =  16 Die Datei ist ein Verzeichnis. Directory wird unter Windows, Linux und macOS unterstützt.
    Archive = VbFileAttribute.vbArchive     '&H20 =  32 Diese Datei ist für das Einbeziehen beim inkrementellen Sicherungsvorgang markiert. Dieses Attribut wird von Windows festgelegt, sobald die
                                            '           Datei geändert wird, und die Sicherungssoftware sollte es beim Verarbeiten der Datei während der inkrementellen Sicherung löschen.
    Device = VbFileAttribute.vbAlias        '&H40 =  64 Für zukünftige Verwendung reserviert.
    Normal = &H80&                          '&H80 = 128 Die Datei ist eine Standarddatei, die über keine speziellen Attribute verfügt. Dieses Attribut ist nur gültig, wenn es allein verwendet wird.
                                            '           Normal wird unter Windows, Linux und macOS unterstützt.
    Temporary = &H100&                      '       256 Die Datei ist temporär. Eine temporäre Datei enthält Daten, die benötigt werden, während eine Anwendung ausgeführt wird, die aber nicht benötigt wird, nachdem die Anwendung beendet ist. Dateisysteme versuchen, für einen schnelleren Zugriff alle Daten im Speicher zu behalten, anstatt diese zurück in den Massenspeicher zu entleeren. Eine temporäre Datei sollte von der Anwendung gelöscht werden, sobald sie nicht mehr benötigt wird.
    SparseFile = &H200&                     '       512 Die Datei ist eine Datei mit geringer Dichte. Dünn besetzte Dateien sind normalerweise große Dateien, deren Daten hauptsächlich aus Nullen bestehen.
    ReparsePoint = &H400&                   '      1024 Die Datei enthält einen Analysepunkt. Dies ist ein Block mit benutzerdefinierten Daten, die mit einer Datei oder einem Verzeichnis verknüpft sind. ReparsePoint wird unter Windows, Linux und macOS unterstützt.
    Compressed = &H800&                     '      2048 Die Datei ist komprimiert.
    Offline = &H1000&                       '      4096 Die Datei ist eine Offlinedatei. Die Daten der Datei sind nicht sofort verfügbar.
    NotContentIndexed = &H2000&             '      8192 Die Datei wird nicht vom Inhaltsindexdienst des Betriebssystems indiziert.
    Encrypted = &H4000&                     '     16384 Die Datei oder das Verzeichnis ist verschlüsselt. Bei einer Datei bedeutet dies, dass alle Daten in der Datei verschlüsselt sind. Bei einem Verzeichnis bedeutet dies, dass neu erstellte Dateien und Verzeichnisse standardmäßig verschlüsselt werden.
    IntegrityStream = &H8000&               '     32768 Die Datei oder das Verzeichnis enthält die Datenintegritätsunterstützung. Wenn dieser Wert auf eine Datei angewendet wird, haben alle Datenstreams in der Datei Integritätsunterstützung. Wenn dieser Wert auf ein Verzeichnis angewendet wird, erhalten alle neuen Dateien und Unterverzeichnisse innerhalb dieses Verzeichnisses standardmäßig Integritätsunterstützung.
    NoScrubData = &H20000                   '    131072 Die Datei oder das Verzeichnis ist von der Datenintegritätsüberprüfung ausgeschlossen. Wenn dieser Wert standardmäßig auf ein Verzeichnis angewendet wird, werden alle neuen Dateien und Unterverzeichnisse innerhalb dieses Verzeichnisses aus der Datenintegrität ausgeschlossen.
End Enum


'old
'Public Enum FileMode
'    fmBinary_Read
'    fmBinary_Write
'    fmBinary_ReadWrite
'    fmText_Read
'    fmText_Write
'End Enum
Public Enum EFileMode
    FileMode_Append
    FileMode_Binary
    FileMode_Input
    FileMode_Output
    FileMode_Random
End Enum
Public Enum EFileAccess
    FileAccess_None
    FileAccess_Read = 1
    FileAccess_Write = 2
    FileAccess_ReadWrite = 3 'Read Or Write is default
End Enum
'Public Enum FileShare
'   FileShare_None = 0
'   FileShare_Read = 1
'   FileShare_Write = 2
'   FileShare_ReadWrite = 3
'   FileShare_Inheritable = 16
'End Enum
Public Enum EFileShare
    FileShare_None = 0
    FileShare_LockRead = 1
    FileShare_LockWrite = 2
    FileShare_LockReadWrite = 3
    FileShare_Shared = 16
End Enum
Private Enum ReadKind
    ReadInput      'Read one thing
    ReadLineInput  'Read one line
    ReadInputFkt   'Read the whole stuff to the end
End Enum
'Oh Shit, 5 * 3 * 4 = 60 verschiedene Möglichkeiten 'oder manche sind siher nicht möglich
Private Const InvalidChars   As String = "\/:*?""<>|" 'bei Dateiname noch "/\:" 'im Pfad darf vorne und hinten kein Leerzeichen sein
Private Const C_DirSepChar   As Long = 92 ' String = "\" ' = Chr(92)
Private Const C_ExtSepChar   As Long = 46 ' String = "." ' = Chr(46)
Private Const C_InvalidChars As String = "*?""<>|" 'bei Dateiname noch "/\:"
Private Const C_3Pt          As String = "..."

Private m_PFN     As String 'the Value
Private m_IsValid As Boolean
Private m_enc     As ETextEncoding
Private m_FNr     As Integer
Private m_FMode   As EFileMode
Private m_FAccess As EFileAccess
Private m_FShare  As EFileShare

' v ############################## v '    Constructors      ' v ############################## v '
Private Sub Class_Initialize()
    m_FAccess = FileAccess_ReadWrite
    m_enc = ETextEncoding.Text_UnicodeEncoding
End Sub
Private Sub Class_Terminate()
    Me.CloseFile
End Sub
'Friend Sub New_(aPath As String, Optional FMode As EFileMode = FileMode_Random, Optional FAccess As FileAccess = FileAccess_ReadWrite, Optional FShare As FileShare = FileShare_None)
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
    m_PFN = other.Value: m_IsValid = other.IsValid
    m_FShare = other.FileShare: m_FMode = other.FileMode: m_FAccess = other.FileAccess
End Sub
Friend Function Clone() As PathFileName
    Set Clone = New PathFileName: Clone.NewC Me 'Value = m_PFN
End Function

'internal ctor
Private Function New_PathFileName(ByVal aPathFileName As String, _
                         Optional ByVal aFileName As String, _
                         Optional ByVal aExt As String) As PathFileName
    Set New_PathFileName = New PathFileName: New_PathFileName.New_ aPathFileName, aFileName, aExt
End Function
''Copy to MNew:
'Public Function PathFileName(ByVal aPathFileName As String, _
'                     Optional ByVal aFileName As String, _
'                     Optional ByVal aExt As String) As PathFileName
'    Set PathFileName = new PathFileName: PathFileName.New_ aPathFileName, aFileName, aExt
'End Function

' ^ ############################## ^ '    Constructors      ' ^ ############################## ^ '

' v ############################## v '  Public Properties   ' v ############################## v '
Public Property Get AttributesX() As EFileAttribute
     AttributesX = GetFileAttributesW(StrPtr(m_PFN))
End Property
Public Property Let AttributesX(ByVal Value As EFileAttribute)
     SetFileAttributesW StrPtr(m_PFN), Value
End Property

Public Function AttributesToStr() As String
    Dim e As EFileAttribute: e = Me.AttributesX
    Dim s As String ': s = ""
    If e = EFileAttribute.None Then s = "None"
    If e And EFileAttribute.ReadOnly Then _
        s = s & IIf(Len(s), " ", "") & "ReadOnly"
    If e And EFileAttribute.Hidden Then _
        s = s & IIf(Len(s), " ", "") & "Hidden"
    If e And EFileAttribute.Volume Then _
        s = s & IIf(Len(s), " ", "") & "Volume"
    If e And EFileAttribute.Directory Then _
        s = s & IIf(Len(s), " ", "") & "Directory"
    If e And EFileAttribute.Archive Then _
        s = s & IIf(Len(s), " ", "") & "Archive"
    If e And EFileAttribute.Device Then _
        s = s & IIf(Len(s), " ", "") & "Device"
    If e And EFileAttribute.Normal Then _
        s = s & IIf(Len(s), " ", "") & "Normal"
    If e And EFileAttribute.Temporary Then _
        s = s & IIf(Len(s), " ", "") & "Temporary"
    If e And EFileAttribute.SparseFile Then _
        s = s & IIf(Len(s), " ", "") & "SparseFile"
    If e And EFileAttribute.ReparsePoint Then _
        s = s & IIf(Len(s), " ", "") & "ReparsePoint"
    If e And EFileAttribute.Compressed Then _
        s = s & IIf(Len(s), " ", "") & "Compressed"
    If e And EFileAttribute.Offline Then _
        s = s & IIf(Len(s), " ", "") & "Offline"
    If e And EFileAttribute.NotContentIndexed Then _
        s = s & IIf(Len(s), " ", "") & "NotContentIndexed"
    If e And EFileAttribute.Encrypted Then _
        s = s & IIf(Len(s), " ", "") & "Encrypted"
    If e And EFileAttribute.IntegrityStream Then _
        s = s & IIf(Len(s), " ", "") & "IntegrityStream"
    If e And EFileAttribute.NoScrubData Then _
        s = s & IIf(Len(s), " ", "") & "NoScrubData"
    AttributesToStr = s
End Function

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

Public Function PathJoin(basepath As PathFileName) As Boolean
    'Path.GetFullPath(Path.Combine("C:\Test1\Test2", "..\Test.pdf")) ' Output: "C:\Test1\Test.pdf"
Try: On Error GoTo Catch
    'Dim pfn_other As String: pfn_other = basepath.Value
    Dim mp() As String: mp = Split(Me.Path, "\")
    Dim op() As String: op = Split(basepath.Path, "\") '.Value
    Dim i As Long, li As Long, ui As Long: ui = UBound(mp)
    Dim j As Long, lj As Long, uj As Long: uj = UBound(op)
    Dim s As String, cj As Long
    For i = 0 To ui
        'shrink mp from below
        'shrink op from above
        s = mp(i)
        If s = ".." Then
            mp(i) = vbNullString
            op(uj) = vbNullString
        End If
        uj = uj - 1
        'what ever comes first
        If uj < 0 Then Exit For
    Next
    s = Join(op, "\") & Join(mp, "\") & Me.FileName
    'Debug.Print s
    Me.New_ s
    PathJoin = True
Catch:
End Function

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

Public Function PathUp() As PathFileName
    Dim p As String: p = m_PFN
    If Right(p, 1) = "\" Then p = Left(p, Len(p) - 1)
    Dim pos As Long: pos = InStrRev(p, "\")
    If pos > 0 Then
        Set PathUp = New_PathFileName(Left(p, pos))
    End If
End Function

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

Public Property Get FileMode() As EFileMode
    FileMode = m_FMode
End Property
Public Property Get FileAccess() As EFileAccess
    FileAccess = m_FAccess
End Property
Public Property Get FileShare() As EFileShare
    FileShare = m_FShare
End Property

' ^ ############################## ^ '  Public Properties   ' ^ ############################## ^ '

' v ####################### v ' Public ReadOnly Boolean Properties ' v ####################### v '
Public Property Get Equals(other As PathFileName) As Boolean
    Equals = Me.Value = other.Value
End Property
Public Property Get IsEqual(other As PathFileName) As Boolean
    IsEqual = Me.Value = other.Value
End Property
Public Property Get IsSame(other As Object) As Boolean
    IsSame = ObjPtr(Me) = ObjPtr(other)
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
Public Property Get IsEOF() As Boolean
    IsEOF = EOF(m_FNr)
End Property

Public Property Get HasBOM() As Boolean
    Dim i As Integer: i = Me.ReadInt16
    HasBOM = i = &HFFFE Or i = &HFEFF
    'If i = &H5B Then
    '    FileIsUnicode = True
    'If i = &HFFFE Or i = &HFEFF Then
    '    i = m_PFN.ReadInt
    '    'If i = &H5B00 Then FileIsUnicode = True
    'End If
    'm_PFN.CloseFile
End Property

' ^ ####################### ^ ' Public ReadOnly Boolean Properties ' ^ ####################### ^ '

' v ############################## v '   Public Functions   ' v ############################## v '
Public Sub OpenFileExplorer()
    Shell "Explorer.exe " & Me.Path, vbNormalFocus
End Sub
'#If VBA7 Then
Public Function Start(Optional ByVal cmd As String = "", Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus) As Long
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
            cmd = IIf(Len(cmd) = 0, m_PFN, m_PFN & " " & cmd)
            ret = Shell(cmd, WindowStyle) ' liefert die TaskID
        Else
            'ein Dokument
            Dim hwnd As LongPtr
            'Debug.Print m_PFN
            ret = ShellExecuteW(hwnd, StrPtr("Open"), StrPtr(Me.Quoted), StrPtr(cmd), StrPtr(Me.Path), WindowStyle)
        End If
    End If
    Start = ret
    Exit Function
Catch:
    ErrHandler "Start"
End Function
    
Public Function StartWait(Optional ByVal cmd As String = "", Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, Optional ByVal WorkDir As String = "") As Long
Try: On Error GoTo Catch
    'startet und wartet = synchron
    'geht jetzt nur bei exe (oder bat?)
    Dim e As String: e = UCase(Me.Extension)
    If Not (e = ".EXE" Or e = ".BAT" Or e = ".COM") Then
        MsgBox "Nur ausführbare Programme (exe) können gestartet werden, Dateiendung: " & Me.Extension
        Exit Function
    End If
    cmd = IIf(Len(cmd) = 0, m_PFN, m_PFN & " " & cmd)
    'Cmd = "C:\Windows\Calc.exe"
    Dim ret As LongPtr: ret = MShell.ShellWait(cmd, WindowStyle, WorkDir)
    'StartWait = ret
    Exit Function
Catch:
    ErrHandler "StartWait", m_PFN, ret
End Function
    
'#Else 'VBA6, VBC
'
'    Public Function Start(Optional ByVal Cmd As String = "", Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus) As Long
'Try:     On Error GoTo Catch
'        'startet ohne zu warten = asynchron
'        Dim ret As Long
'        If Me.FileName = "" Then
'            Dim p As PathFileName: Set p = Me
'            Do While Not p Is Nothing
'                If p.PathExists Then Exit Do
'                'If Not p.PathExists Then
'                Set p = p.PathUp
'            Loop
'            If Not p Is Nothing Then
'                ret = Shell("Explorer.exe " & p.Path, vbNormalFocus)
'            End If
'            Exit Function
'        Else
'            Dim e As String: e = UCase(Me.Extension)
'            If e = "EXE" Or e = "BAT" Then
'                'ein ausführbares programm
'                Cmd = IIf(Len(Cmd) = 0, m_PFN, m_PFN & " " & Cmd)
'                ret = Shell(Cmd, WindowStyle) ' liefert die TaskID
'            Else
'                'ein Dokument
'                Dim hwnd As Long
'                'Debug.Print m_PFN
'                ret = ShellExecuteW(hwnd, StrPtr("Open"), StrPtr(Me.Quoted), StrPtr(Cmd), StrPtr(Me.Path), WindowStyle)
'            End If
'        End If
'        Start = ret
'        Exit Function
'Catch:
'        ErrHandler "Start"
'    End Function
'
'    Public Function StartWait(Optional ByVal Cmd As String = "", Optional ByVal WindowStyle As VbAppWinStyle = vbNormalFocus, Optional ByVal WorkDir As String) As Long
'Try:     On Error GoTo Catch
'        'startet und wartet = synchron
'        'geht jetzt nur bei exe (oder bat?)
'        Dim e As String: e = UCase(Me.Extension)
'        If Not (e = ".EXE" Or e = ".BAT" Or e = ".COM") Then
'            MsgBox "Start executable programs only (like .exe, .bat or .com), your file extension: " & Me.Extension
'            Exit Function
'        End If
'        Dim ret As Long
'        Cmd = IIf(Len(Cmd) = 0, m_PFN, m_PFN & " " & Cmd)
'        ret = MShell.ShellWait(Cmd, WindowStyle, WorkDir)
'        'StartWait = ret
'        Exit Function
'Catch:
'        ErrHandler "StartWait"
'    End Function
'
'#End If

Public Function RemoveQuotes(ByVal aStrPFN As String) As String
    If Len(aStrPFN) = 0 Then Exit Function
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
    'returns the pointer to the filename-string, if there is no value, a buffer
    'will be created an the pointer to it will be returned
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

' v ############################## v '    Path related functions    ' v ############################## v '
'was macht PathCreate?
'erzeugt es ein neues Verzeichnis im Laufwerks-Root oder im aktuellen Verzeichnis?
'
Public Function PathCreate(Optional ByVal sNewDir As String) As Boolean
Try: On Error GoTo Catch
    'ein neues Verzeichnis mit Unterverzeichnissen anlegen
    If Len(sNewDir) = 0 Then sNewDir = Me.Path
    Call MakeValid(sNewDir)
    Dim hr As Long
    hr = MakeSureDirectoryPathExists(StrPtr(StrConv(sNewDir, vbFromUnicode)))
    If hr = 0 Then GoTo Catch
    'Makesuredirectorypathexits
'    Dim NewDir As PathFileName: Set NewDir = New_PathFileName(sNewDir)
'    'zuerst in das jeweilige Laufwerk wechseln
'    Dim sd As String: sd = NewDir.Drive()
'    Call ChDir(sd)
'    Dim i As Long
'    For i = 0 To NewDir.PathCount - 1
'        sd = NewDir.PathI(i)
'        'das Verzeichnis anlegen
'        Call MkDir(sd)
'        'gleich ins verzeichnis wechseln, damit das nächste
'        'Unterverzeichnis angelegt werden kann
'        Call ChDir(sd)
'    Next
    PathCreate = True
    Exit Function
Catch:
    ErrHandler "PathCreate", "Could not create path: " & vbCrLf & sNewDir
End Function

Public Function PathAdd(Optional ByVal NewSubDir As String) As Boolean
    Me.Path = Me.Path & NewSubDir & "\"
    PathAdd = PathCreate()
End Function
'a path is a stack so we need the words Peek, Push, Pop
Public Function PathPush(ByVal NewSubDir As String) As Boolean
    PathPush = PathAdd(NewSubDir)
End Function
Public Function PathPop() As String
    'den obersten Pfad extrahieren zurückgeben und wegnehmen aber nicht direk physikalisch von der Platte löschen!
    
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
    Set AddToFileNameR = New_PathFileName(Me.Path & Me.FileNameOnly & addstr & IIf(Len(e), "." & e, ""))
End Function
Public Function AddToFileNameL(ByVal addstr As String) As PathFileName
    Dim e As String: e = Me.Extension 'Änderung 08.01.2016
    Set AddToFileNameL = New_PathFileName(Me.Path & addstr & Me.FileNameOnly & IIf(Len(e), "." & e, ""))
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
        Exists = CBool(GetAttr(p) And vbDirectory)
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
    PathExists = CBool(GetAttr(p) And vbDirectory)
    'If Not IsInIDE Then
    '    PathExists = CBool(GetAttr(p) And vbDirectory)
    'End If
    On Error GoTo 0
End Function

Public Property Get Encoding() As ETextEncoding
    Encoding = m_enc 'Text_UnicodeEncoding
End Property
Public Property Let Encoding(ByVal Value As ETextEncoding)
    m_enc = Value
End Property
    

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
        Set newPFN = New_PathFileName(newPathFileName.Path & Me.FileName)
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
    Dim rv As VbMsgBoxResult: rv = ErrHandler("CopyRetryLoop", m_PFN & vbCrLf & newPathFileName.Value, False, , , True)
    ' MessErrorRetry(Me, "CopyTo", m_PFN & vbCrLf & newPathFileName.Value)
    If rv = vbRetry Then
        Resume Try 'the user has the choice to stop it anytime
    Else
        Set CopyRetryLoop = Nothing
        Exit Function
    End If
End Function

Public Function MoveTo(PFNdst As PathFileName) As PathFileName
    'move = nach Dst kopieren und Src löschen
    'Move = verschieben, im Quellordner ist die Datei weg
    'Copy = Kopieren,    im Quellordner ist die Datei noch da
Try: On Error GoTo Catch
    Set MoveTo = Me.CopyTo(PFNdst)
    If MoveTo Is Nothing Then Exit Function
    Me.Delete
    Exit Function
Catch:
    ErrHandler "MoveTo", m_PFN & vbCrLf & PFNdst.Value
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
    Dim RenFile As PathFileName: Set RenFile = New_PathFileName(ToPFN)
    Dim cmd     As String
    If Not RenFile.Exists Then
        'Wird nicht renamed wenn es die Datei schon gibt
        cmd = Chr(34) & Me.Value & Chr(34) & " " & Chr(34) & RenFile.Value & Chr(34)
        Call ShellWait("cmd /c ren " & cmd, VbAppWinStyle.vbHide)
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

Public Property Get Length() As Long
    'Returns the length or size of the file in bytes
    If m_FNr Then Length = LOF(m_FNr) Else Length = FileLen(m_PFN)
End Property

Public Property Let pos(ByVal Value As Long)
    Seek m_FNr, Value
End Property

Public Property Get pos() As Long
    pos = Seek(m_FNr)
End Property
'Functions for Writing to and Reading from files
'Public Function OpenFile(fm As EFileMode) As Integer
'    m_fm = fm
'Try: On Error GoTo Catch
'    If m_FNr <> 0 Then OpenFile = m_FNr: Exit Function
'    m_FNr = FreeFile
'    Select Case m_fm
'    Case fmBinary_Read
'
'        Open m_PFN For Binary Access Read As m_FNr
'
'    Case fmBinary_Write
'
'        Open m_PFN For Binary Access Write As m_FNr
'
'    Case fmBinary_ReadWrite
'
'        Open m_PFN For Binary As m_FNr
'
'    Case fmText_Read
'
'        Open m_PFN For Input As m_FNr
'
'    Case fmText_Write
'
'        Open m_PFN For Output As m_FNr
'
'    End Select
'    OpenFile = m_FNr
'    Exit Function
'Catch:
'    ErrHandler "OpenFile", m_PFN
'End Function

'Friend Sub New_(aPath As String, Optional FMode As EFileMode = EFileMode.FileMode_Random, Optional FAccess As EFileAccess = FileAccess_ReadWrite, Optional FShare As FileShare = FileShare_None)
'Private Function OOpen(ByVal aFMode As EFileMode, ByVal aFAccess As FileAccess, ByVal aFShare As FileShare) As Boolean
Public Function OpenFile(Optional ByVal fm As EFileMode = EFileMode.FileMode_Random, _
                         Optional ByVal fa As EFileAccess = EFileAccess.FileAccess_ReadWrite, _
                         Optional ByVal fs As EFileShare = EFileShare.FileShare_None) As Integer
Try: On Error GoTo Catch
    'aFMode = mFMode
    'aFAccess = mFAccess
    'aFShare = mFShare
    If m_FNr <> 0 Then OpenFile = m_FNr: Exit Function
    m_FNr = FreeFile
    Dim pfn As String: pfn = Me.Value
    m_FMode = fm
    m_FAccess = fa
    m_FShare = fs
    Select Case fm
    Case EFileMode.FileMode_Append
        Select Case fa
        Case EFileAccess.FileAccess_None
            Select Case fs
            Case EFileShare.FileShare_None:          Open pfn For Append As m_FNr
            Case EFileShare.FileShare_LockRead:      Open pfn For Append Lock Read As m_FNr
            Case EFileShare.FileShare_LockWrite:     Open pfn For Append Lock Write As m_FNr
            Case EFileShare.FileShare_LockReadWrite: Open pfn For Append Lock Read Write As m_FNr
            Case EFileShare.FileShare_Shared:        Open pfn For Append Shared As m_FNr
            End Select
        'Case EFileAccess.FileAccess_Read 'Not Possible
        '    Select Case aFShare
        '    Case FileShare_None:          Open mPath For Append Access Read As m_FNr
        '    Case FileShare_LockRead:      Open mPath For Append Access Read Lock Read As m_FNr
        '    Case FileShare_LockWrite:     Open mPath For Append Access Read Lock Write As m_FNr
        '    Case FileShare_LockReadWrite: Open mPath For Append Access Read Lock Read Write As m_FNr
        '    Case FileShare_Shared:        Open mPath For Append Access Read Shared As m_FNr
        '    End Select
        Case EFileAccess.FileAccess_Write
            Select Case fs
            Case EFileShare.FileShare_None:          Open pfn For Append Access Write As m_FNr
            Case EFileShare.FileShare_LockRead:      Open pfn For Append Access Write Lock Read As m_FNr
            Case EFileShare.FileShare_LockWrite:     Open pfn For Append Access Write Lock Write As m_FNr
            Case EFileShare.FileShare_LockReadWrite: Open pfn For Append Access Write Lock Read Write As m_FNr
            Case EFileShare.FileShare_Shared:        Open pfn For Append Access Write Shared As m_FNr
            End Select
        Case EFileAccess.FileAccess_ReadWrite
            Select Case fs
            Case EFileShare.FileShare_None:          Open pfn For Append Access Read Write As m_FNr
            Case EFileShare.FileShare_LockRead:      Open pfn For Append Access Read Write Lock Read As m_FNr
            Case EFileShare.FileShare_LockWrite:     Open pfn For Append Access Read Write Lock Write As m_FNr
            Case EFileShare.FileShare_LockReadWrite: Open pfn For Append Access Read Write Lock Read Write As m_FNr
            Case EFileShare.FileShare_Shared:        Open pfn For Append Access Read Write Shared As m_FNr
            End Select
        End Select
    Case EFileMode.FileMode_Binary
        Select Case fa
        Case EFileAccess.FileAccess_None
            Select Case fs
            Case EFileShare.FileShare_None:          Open pfn For Binary As m_FNr
            Case EFileShare.FileShare_LockRead:      Open pfn For Binary Lock Read As m_FNr
            Case EFileShare.FileShare_LockWrite:     Open pfn For Binary Lock Write As m_FNr
            Case EFileShare.FileShare_LockReadWrite: Open pfn For Binary Lock Read Write As m_FNr
            Case EFileShare.FileShare_Shared:        Open pfn For Binary Shared As m_FNr
            End Select
        Case EFileAccess.FileAccess_Read
            Select Case fs
            Case EFileShare.FileShare_None:          Open pfn For Binary Access Read As m_FNr
            Case EFileShare.FileShare_LockRead:      Open pfn For Binary Access Read Lock Read As m_FNr
            Case EFileShare.FileShare_LockWrite:     Open pfn For Binary Access Read Lock Write As m_FNr
            Case EFileShare.FileShare_LockReadWrite: Open pfn For Binary Access Read Lock Read Write As m_FNr
            Case EFileShare.FileShare_Shared:        Open pfn For Binary Access Read Shared As m_FNr
            End Select
        Case EFileAccess.FileAccess_Write
            Select Case fs
            Case EFileShare.FileShare_None:          Open pfn For Binary Access Write As m_FNr
            Case EFileShare.FileShare_LockRead:      Open pfn For Binary Access Write Lock Read As m_FNr
            Case EFileShare.FileShare_LockWrite:     Open pfn For Binary Access Write Lock Write As m_FNr
            Case EFileShare.FileShare_LockReadWrite: Open pfn For Binary Access Write Lock Read Write As m_FNr
            Case EFileShare.FileShare_Shared:        Open pfn For Binary Access Write Shared As m_FNr
            End Select
        Case EFileAccess.FileAccess_ReadWrite
            Select Case fs
            Case EFileShare.FileShare_None:          Open pfn For Binary Access Read Write As m_FNr
            Case EFileShare.FileShare_LockRead:      Open pfn For Binary Access Read Write Lock Read As m_FNr
            Case EFileShare.FileShare_LockWrite:     Open pfn For Binary Access Read Write Lock Write As m_FNr
            Case EFileShare.FileShare_LockReadWrite: Open pfn For Binary Access Read Write Lock Read Write As m_FNr
            Case EFileShare.FileShare_Shared:        Open pfn For Binary Access Read Write Shared As m_FNr
            End Select
        End Select
    Case EFileMode.FileMode_Input
        Select Case fa
        Case EFileAccess.FileAccess_None
            Select Case fs
            Case EFileShare.FileShare_None:          Open pfn For Input As m_FNr
            Case EFileShare.FileShare_LockRead:      Open pfn For Input Lock Read As m_FNr
            Case EFileShare.FileShare_LockWrite:     Open pfn For Input Lock Write As m_FNr
            Case EFileShare.FileShare_LockReadWrite: Open pfn For Input Lock Read Write As m_FNr
            Case EFileShare.FileShare_Shared:        Open pfn For Input Shared As m_FNr
            End Select
        Case EFileAccess.FileAccess_Read, EFileAccess.FileAccess_ReadWrite
            Select Case fs
            Case EFileShare.FileShare_None:          Open pfn For Input Access Read As m_FNr
            Case EFileShare.FileShare_LockRead:      Open pfn For Input Access Read Lock Read As m_FNr
            Case EFileShare.FileShare_LockWrite:     Open pfn For Input Access Read Lock Write As m_FNr
            Case EFileShare.FileShare_LockReadWrite: Open pfn For Input Access Read Lock Read Write As m_FNr
            Case EFileShare.FileShare_Shared:        Open pfn For Input Access Read Shared As m_FNr
            End Select
        'Case EFileAccess 'Not Possible
        '    Select Case fs
        '    Case FileShare_None:          Open mPath For Input Access Write As m_FNr
        '    Case FileShare_LockRead:      Open mPath For Input Access Write Lock Read As m_FNr
        '    Case FileShare_LockWrite:     Open mPath For Input Access Write Lock Write As m_FNr
        '    Case FileShare_LockReadWrite: Open mPath For Input Access Write Lock Read Write As m_FNr
        '    Case FileShare_Shared:        Open mPath For Input Access Write Shared As m_FNr
        '    End Select
        'Case EFileAccess 'Not Possible
        '    Select Case fs
        '    Case FileShare_None:          Open mPath For Input Access Read Write As m_FNr
        '    Case FileShare_LockRead:      Open mPath For Input Access Read Write Lock Read As m_FNr
        '    Case FileShare_LockWrite:     Open mPath For Input Access Read Write Lock Write As m_FNr
        '    Case FileShare_LockReadWrite: Open mPath For Input Access Read Write Lock Read Write As m_FNr
        '    Case FileShare_Shared:        Open mPath For Input Access Read Write Shared As m_FNr
        '    End Select
        End Select
    Case EFileMode.FileMode_Output
        Select Case fa
        Case EFileAccess.FileAccess_None
            Select Case fs
            Case EFileShare.FileShare_None:          Open pfn For Output As m_FNr
            Case EFileShare.FileShare_LockRead:      Open pfn For Output Lock Read As m_FNr
            Case EFileShare.FileShare_LockWrite:     Open pfn For Output Lock Write As m_FNr
            Case EFileShare.FileShare_LockReadWrite: Open pfn For Output Lock Read Write As m_FNr
            Case EFileShare.FileShare_Shared:        Open pfn For Output Shared As m_FNr
            End Select
'        Case EFileAccess
'            Select Case fs 'Not Possible
'            Case FileShare_None:          Open mPath For Output Access Read As #mFNr
'            Case FileShare_LockRead:      Open mPath For Output Access Read Lock Read As #mFNr
'            Case FileShare_LockWrite:     Open mPath For Output Access Read Lock Write As #mFNr
'            Case FileShare_LockReadWrite: Open mPath For Output Access Read Lock Read Write As #mFNr
'            Case FileShare_Shared:        Open mPath For Output Access Read Shared As #mFNr
'            End Select
        Case EFileAccess.FileAccess_Write
            Select Case fs
            Case EFileShare.FileShare_None:          Open pfn For Output Access Write As m_FNr
            Case EFileShare.FileShare_LockRead:      Open pfn For Output Access Write Lock Read As m_FNr
            Case EFileShare.FileShare_LockWrite:     Open pfn For Output Access Write Lock Write As m_FNr
            Case EFileShare.FileShare_LockReadWrite: Open pfn For Output Access Write Lock Read Write As m_FNr
            Case EFileShare.FileShare_Shared:        Open pfn For Output Access Write Shared As m_FNr
            End Select
            '    Case FileAccess_ReadWrite
            '      Select Case fs 'Not Possible
            '      Case FileShare_None:          Open mPath For Output Access Read Write As m_FNr
            '      Case FileShare_LockRead:      Open mPath For Output Access Read Write Lock Read As m_FNr
            '      Case FileShare_LockWrite:     Open mPath For Output Access Read Write Lock Write As m_FNr
            '      Case FileShare_LockReadWrite: Open mPath For Output Access Read Write Lock Read Write As m_FNr
            '      Case FileShare_Shared:        Open mPath For Output Access Read Write Shared As m_FNr
            '      End Select
        End Select
    Case EFileMode.FileMode_Random
        Select Case fa
        Case EFileAccess.FileAccess_None
            Select Case fs
            Case EFileShare.FileShare_None:          Open pfn For Random As m_FNr
            Case EFileShare.FileShare_LockRead:      Open pfn For Random Lock Read As m_FNr
            Case EFileShare.FileShare_LockWrite:     Open pfn For Random Lock Write As m_FNr
            Case EFileShare.FileShare_LockReadWrite: Open pfn For Random Lock Read Write As m_FNr
            Case EFileShare.FileShare_Shared:        Open pfn For Random Shared As m_FNr
            End Select
        Case EFileAccess.FileAccess_Read
            Select Case fs
            Case EFileShare.FileShare_None:          Open pfn For Random Access Read As m_FNr
            Case EFileShare.FileShare_LockRead:      Open pfn For Random Access Read Lock Read As m_FNr
            Case EFileShare.FileShare_LockWrite:     Open pfn For Random Access Read Lock Write As m_FNr
            Case EFileShare.FileShare_LockReadWrite: Open pfn For Random Access Read Lock Read Write As m_FNr
            Case EFileShare.FileShare_Shared:        Open pfn For Random Access Read Shared As m_FNr
            End Select
        Case EFileAccess.FileAccess_Write
            Select Case fs
            Case EFileShare.FileShare_None:          Open pfn For Random Access Write As m_FNr
            Case EFileShare.FileShare_LockRead:      Open pfn For Random Access Write Lock Read As m_FNr
            Case EFileShare.FileShare_LockWrite:     Open pfn For Random Access Write Lock Write As m_FNr
            Case EFileShare.FileShare_LockReadWrite: Open pfn For Random Access Write Lock Read Write As m_FNr
            Case EFileShare.FileShare_Shared:        Open pfn For Random Access Write Shared As m_FNr
            End Select
        Case EFileAccess.FileAccess_ReadWrite
            Select Case fs
            Case EFileShare.FileShare_None:          Open pfn For Random Access Read As m_FNr
            Case EFileShare.FileShare_LockRead:      Open pfn For Random Access Read Write Lock Read As m_FNr
            Case EFileShare.FileShare_LockWrite:     Open pfn For Random Access Read Write Lock Write As m_FNr
            Case EFileShare.FileShare_LockReadWrite: Open pfn For Random Access Read Write Lock Read Write As m_FNr
            Case EFileShare.FileShare_Shared:        Open pfn For Random Access Read Write Shared As m_FNr
            End Select
        End Select
    End Select
    OpenFile = m_FNr
    Exit Function
Catch:
    MsgBox "Error: " & CStr(Err.Number) & " in PFNStream::OOpen: " & Err.Description
End Function

Public Function ReadAllText() As String
Try: On Error GoTo Catch
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, EFileAccess.FileAccess_Read, EFileShare.FileShare_Shared  ' fmBinary_Read
    If m_FNr = 0 Then Exit Function
    Dim bom As EByteOrderMark: bom = ReadBOM
    Dim s As String
    Dim Buffer() As Byte
    Select Case m_FMode
    Case EFileMode.FileMode_Binary
        Select Case bom
        Case EByteOrderMark.bom_UTF_16_LE
            ReDim Buffer(0 To LOF(m_FNr) - Me.pos + 1)
            Get m_FNr, , Buffer
            ReadAllText = Buffer
        Case EByteOrderMark.bom_UTF_16_BE
            ReDim Buffer(0 To LOF(m_FNr) - Me.pos + 1)
            Get m_FNr, , Buffer
            ReadAllText = Buffer
            MPtr.String_Rotate2 ReadAllText
        Case EByteOrderMark.bom_UTF_32_LE
            s = Space(LOF(m_FNr) - Me.pos + 1)
            Get m_FNr, , s
            ReadAllText = StrConv(s, vbFromUnicode)
        Case EByteOrderMark.bom_UTF_32_BE
            s = Space(LOF(m_FNr) - Me.pos + 1)
            Get m_FNr, , s
            ReadAllText = StrConv(s, vbFromUnicode)
            MPtr.String_Rotate2 ReadAllText
        Case EByteOrderMark.bom_UTF_7
            ReDim Buffer(0 To LOF(m_FNr) - Me.pos + 1)
            Get m_FNr, , Buffer
            ReadAllText = Buffer
        Case EByteOrderMark.bom_UTF_8
            ReDim Buffer(0 To LOF(m_FNr) - Me.pos + 1)
            Get m_FNr, , Buffer
            ReadAllText = MString.ConvertFromUTF8(Buffer)
        'Case EByteOrderMark.bom_UTF_8
        
        Case Else
            'there is not bom, maybe xml file?
            Select Case m_enc
            Case ETextEncoding.Text_UTF8Encoding
                ReDim Buffer(0 To LOF(m_FNr) - Me.pos + 1)
                Get m_FNr, , Buffer
                s = MString.Encoding_GetString(m_enc, Buffer)
            Case Else
                s = Space(LOF(m_FNr) - Me.pos + 1)
                Get m_FNr, , s
            End Select
            
            ReadAllText = s
        End Select
            
    Case EFileMode.FileMode_Input
        Input #m_FNr, s
        ReadAllText = s
    End Select
    Me.CloseFile
    Exit Function
Catch:
    ErrHandler "ReadAllText", m_PFN
End Function

'https://de.wikipedia.org/wiki/Byte_Order_Mark
'Public Enum EByteOrderMark
'    bom_UTF_8 = &HBFBBEF        '                     '     239 187 191     '  ï»¿           ' [4]
'    bom_UTF_16_BE = &HFFFE&     ' Big Endian Motorola '         254 255     '   þÿ
'    bom_UTF_16_LE = &HFEFF&     ' little endian Intel '         255 254     '   ÿþ
'    bom_UTF_32_BE = &HFFFE0000  ' Big Endian Motorola '   0   0 254 255     ' ??þÿ
'    bom_UTF_32_LE = &HFEFF      ' little endian Intel ' 255 254   0   0     ' ÿþ??
'    bom_UTF_7 = &H762F2B        '                     '      43  47 118
'                                ' und ein Zeichen aus: [ 56 | 57 | 43 | 47 ]
'                                ' und ein Zeichen aus: [ 38 | 39 | 2B | 2F ]           ' [5]
'                                ' + / v und ein Zeichen aus:  [  8 |  9 |  + |  / ]
'    bom_UTF_1 = &H4C64F7        '                     '     247 100  76     ' ÷dL
'    bom_UTF_EBCDIC = &H736673DD '                     ' 221 115 102 115     ' Ýsfs
'    bom_SCSU = &HFFFE0E         '                     '      14 254 255     ' ?þÿ            ' [6]
'                                ' (von anderen möglichen Bytefolgen wird abgeraten)
'    bom_BOCU_1 = &H28EEFB       '                     '     251 238  40
'                                ' optional gefolgt von FF                              ' [7]
'                                ' optional gefolgt von 255     ûî
'                                ' optional gefolgt von          ÿ
'    bom_GB_18030 = &H33953184   '               ' 132  49 149  51     ' „1•3
'End Enum

'Public Enum EByteOrderMark
'Byte-Nr                  4  3  2  1
'    bom_UTF_32_BE  = &H FF FE 00 00 ' Big Endian Motorola '   0   0 254 255     ' ??þÿ
'    bom_UTF_32_LE  = &H 00 00 FE FF ' little endian Intel ' 255 254   0   0     ' ÿþ??
'    bom_GB_18030   = &H 33 95 31 84 '                     ' 132  49 149  51     ' „1•3
'    bom_UTF_EBCDIC = &H 73 66 73 DD '                     ' 221 115 102 115     ' Ýsfs

'    bom_SCSU       = &H    FF FE 0E '                     '      14 254 255     ' ?þÿ      ' [6]
'                                    ' (von anderen möglichen Bytefolgen wird abgeraten)
'    bom_UTF_8      = &H    BF BB EF '                     '     239 187 191     '  ï»¿     ' [4]
'    bom_UTF_1      = &H    4C 64 F7 '                     '     247 100  76     ' ÷dL

'    bom_UTF_7      = &H    76 2F 2B '                     '      43  47 118
'                                    ' und ein Zeichen aus: [ 56 | 57 | 43 | 47 ]
'                                    ' und ein Zeichen aus: [ 38 | 39 | 2B | 2F ]           ' [5]
'                                    ' + / v und ein Zeichen aus:  [  8 |  9 |  + |  / ]

'    bom_BOCU_1     = &H    28 EE FB '                     '     251 238  40
'                                    ' optional gefolgt von FF                              ' [7]
'                                    ' optional gefolgt von 255     ûî
'                                    ' optional gefolgt von          ÿ

'    bom_UTF_16_BE  = &H       FF FE ' Big Endian Motorola '         254 255     '   þÿ
'    bom_UTF_16_LE  = &H       FE FF ' little endian Intel '         255 254     '   ÿþ

'End Enum

Private Function ReadBOM() As EByteOrderMark
    
    If m_FNr = 0 Then Me.OpenFile EFileMode.FileMode_Binary, FileAccess_Read, FileShare_Shared
    Dim e As EByteOrderMark: Get m_FNr, 1, e 'read 4 bytes from position 0
    e = MString.EByteOrderMark_Parse(e)
    'jetzt seek setzen
    Dim fpos As Long
    Select Case e
    Case EByteOrderMark.bom_UTF_32_BE  ': nix machen
    Case EByteOrderMark.bom_UTF_32_LE  ': nix machen
    Case EByteOrderMark.bom_GB_18030   ': nix machen
    Case EByteOrderMark.bom_UTF_EBCDIC ': nix machen
    Case EByteOrderMark.bom_UTF_7      ': nix machen
    Case EByteOrderMark.bom_SCSU:      fpos = 4
    Case EByteOrderMark.bom_UTF_8:     fpos = 4
    Case EByteOrderMark.bom_UTF_1:     fpos = 4
    Case EByteOrderMark.bom_BOCU_1:    fpos = 4 'maybe fpos=4
    Case EByteOrderMark.bom_UTF_16_BE: fpos = 3
    Case EByteOrderMark.bom_UTF_16_LE: fpos = 3
    Case Else: fpos = 1
    End Select
    Me.pos = fpos
    ReadBOM = e
End Function

'private function Bom_Equals(
'Public Function ReadAllStrW() As String
'Try: On Error GoTo Catch
'    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary ', FileAccess.FileAccess_Read ' fmBinary_Read
'    Dim aStr() As Byte
'    Select Case m_FMode
'    Case EFileMode.FileMode_Binary 'fmBinary_Read
'        ReDim aStr(0 To LOF(m_FNr))
'        Get #m_FNr, , aStr
'    Case EFileMode.FileMode_Input 'fmText_Read
'        Input #m_FNr, aStr
'    End Select
'    ReadAllStrW = aStr
'    Exit Function
'Catch:
'    ErrHandler "ReadAllStrW", m_PFN
'End Function

Public Function ReadLine() As String
    If m_FNr = 0 Then
        Me.OpenFile EFileMode.FileMode_Input, FileAccess_Read
    End If
    If m_FMode = EFileMode.FileMode_Input Then
        Line Input #m_FNr, ReadLine
    Else
        'wenn binary dann bis zum, nächsten vbcr, vblf oder vbcrlf lesen
    End If
End Function

Public Function WriteLine(aLine As String) As Boolean
Try: On Error GoTo Catch
    If m_FNr = 0 Then
        Me.OpenFile EFileMode.FileMode_Output, EFileAccess.FileAccess_Write
    End If
    If m_FMode = EFileMode.FileMode_Output Then
        Print #m_FNr, aLine
    Else
        'wenn binary dann bis zum, nächsten vbcr, vblf oder vbcrlf lesen
    End If
Catch:
End Function

Public Function WriteLineW(ByVal aLine As String) As Boolean
Try: On Error GoTo Catch
    If m_FNr = 0 Then
        Me.OpenFile EFileMode.FileMode_Binary, EFileAccess.FileAccess_Write
    End If
    If m_FMode = EFileMode.FileMode_Binary Then
        Dim Buffer() As Byte: Buffer = aLine & vbCrLf
        'Write #m_FNr, aLine & vbCrLf
        Put #m_FNr, , Buffer
    Else
        'works only with EFileMode.FileMode_Binary
    End If
Catch:
End Function

Public Function TryReadAllLines(lines_out() As String) As Boolean
    Dim s As String: s = Me.ReadAllText
    If Len(s) = 0 Then Exit Function
    lines_out = Split(s, vbCrLf)
    TryReadAllLines = True
End Function

Public Sub CloseFile()
    If m_FNr = 0 Then Exit Sub
Try: On Error GoTo Catch
    Close m_FNr
    m_FNr = 0
    Exit Sub
Catch:
    ErrHandler "CloseFile", m_PFN
End Sub

' v ############################## v '    BinaryReader/Writer Functions    ' v ############################## v '
Public Function ReadByte() As Byte
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Get m_FNr, , ReadByte
End Function
Public Sub WriteByte(ByVal Value As Byte)
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Put m_FNr, , Value
End Sub

Public Sub WriteBytes(ByteArray() As Byte)
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Put m_FNr, , ByteArray
End Sub

Public Function ReadChrW() As Integer
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Get m_FNr, , ReadChrW
End Function
Public Sub WriteChrW(ByVal aCharW As Integer)
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Put m_FNr, , aCharW
End Sub

Public Function ReadInt16() As Integer
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Get m_FNr, , ReadInt16
End Function
Public Sub WriteInt16(ByVal Value As Integer)
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Put m_FNr, , Value
End Sub

Public Function ReadInt32() As Long
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Get m_FNr, , ReadInt32
End Function
Public Sub WriteInt32(ByVal Value As Long)
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Put m_FNr, , Value
End Sub

Public Function ReadCur() As Currency
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Get m_FNr, , ReadCur
End Function
Public Sub WriteCur(ByVal Value As Currency)
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Put m_FNr, , Value
End Sub

Public Function ReadSingle() As Single
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Get m_FNr, , ReadSingle
End Function
Public Sub WriteSingle(ByVal Value As Single)
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Put m_FNr, , Value
End Sub

Public Function ReadDouble() As Double
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Get m_FNr, , ReadDouble
End Function
Public Sub WriteDouble(ByVal Value As Double)
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Put m_FNr, , Value
End Sub

Public Function ReadDate() As Date
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Get m_FNr, , ReadDate
End Function
Public Sub WriteDate(ByVal Value As Date)
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Put m_FNr, , Value
End Sub

Public Function ReadVariant() As Variant
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Get m_FNr, , ReadVariant
End Function
Public Sub WriteVariant(ByVal Value As Variant)
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Put m_FNr, , Value
End Sub

Public Function ReadStr(ByVal BytLen As Long) As String
    'Reads Ansi String
    ReadStr = String$(BytLen, vbNullChar)
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, FileAccess_ReadWrite
    Get m_FNr, , ReadStr
End Function
'Public Sub WriteStr(ByVal Value As String)
'    If m_FNr = 0 Then OpenFile FileMode_Binary, FileAccess_ReadWrite
'    Put m_FNr, , Value
'End Sub

Public Function WriteStr(Value As String) As Boolean
    'Writes Ansi String
Try: On Error GoTo Catch
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary ', FileAccess.FileAccess_Write
    Select Case m_FMode
    Case EFileMode.FileMode_Append
        'mit Put Err54 falscher Dateimodus
        'Write #m_FNr, aStr
        Print #m_FNr, Value
    Case EFileMode.FileMode_Binary
        Put #m_FNr, , Value
    Case EFileMode.FileMode_Output
        Write #m_FNr, Value
    End Select
    WriteStr = True
    Exit Function
Catch:
    ErrHandler "WriteStr", m_PFN
End Function

Public Function ReadStrW(ByVal BytLen As Long) As String
    'Reads WideChar String
    ReDim Buffer(BytLen) As Byte
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, EFileAccess.FileAccess_ReadWrite
    Get m_FNr, , Buffer
    ReadStrW = Buffer
End Function
Public Sub WriteStrW(ByVal Value As String)
    'Writes WideChar String
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, EFileAccess.FileAccess_ReadWrite
    Dim Buffer() As Byte: Buffer = Value
    Put m_FNr, , Buffer
End Sub

Public Sub ReadMem(ByVal pDst As LongPtr, ByVal BytLen As Long)
    ReDim Buffer(0 To BytLen) As Byte
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, EFileAccess.FileAccess_ReadWrite
    Get m_FNr, , Buffer
    RtlMoveMemory ByVal pDst, Buffer(0), BytLen
End Sub
Public Sub WriteMem(ByVal pSrc As LongPtr, ByVal BytLen As Long)
    ReDim Buffer(0 To BytLen) As Byte
    RtlMoveMemory Buffer(0), ByVal pSrc, BytLen
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, EFileAccess.FileAccess_ReadWrite
    Put m_FNr, , Buffer
End Sub

Public Sub ReadAllBuffer(Buffer() As Byte)
Try: On Error GoTo Catch
    If m_FNr = 0 Then OpenFile EFileMode.FileMode_Binary, EFileAccess.FileAccess_ReadWrite
    ReDim Buffer(0 To LOF(m_FNr) - 1)
    Get m_FNr, , Buffer
    GoTo Finally
Catch:
    If Err Then ErrHandler "ReadAllBuffer"
Finally:
    Me.CloseFile
End Sub
' ^ ############################## ^ '    BinaryReader/Writer Functions    ' ^ ############################## ^ '

' ^ ############################## ^ '   Public Functions   ' ^ ############################## ^ '

' v ############################## v '   Private Parsers    ' v ############################## v '
Private Function Get_Drive(PFN_in As String) As String
'  "\\?\",
    'can be either like "C:\" or like "\\Server\Data\" (UNC);
    'a local drive contains ":", a network-drive contains "\\" at the beginning;
    'and yes, a linked network drive will be treated here as a local drive
    'it will always be closed by "\"
    Dim s As String
    Dim pos2 As Long, pos As Long: pos = InStr(1, PFN_in, ":")
    If 0 < pos And pos < 3 Then
        s = Left$(PFN_in, pos) & "\"
    ElseIf Left$(PFN_in, 2) = "\\" Then
        pos = InStr(3, PFN_in, "\")
        If pos = 0 Then
            If Right$(PFN_in, 1) <> "\" Then
                s = PFN_in & "\"
            End If
        ElseIf pos > 2 Then
            'read to the next "\"
            'pos2 = pos
            pos2 = InStr(pos + 1, PFN_in, "\")
            If pos2 > 0 Then pos = Max(pos, pos2)
            s = Left$(PFN_in, pos)
        End If
    End If
    Get_Drive = s
End Function
Private Function Max(V1, V2)
    If V1 > V2 Then Max = V1 Else Max = V2
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
'Die funktion MakeValid sollte die Bestandteile des Pfades getrennt voneinander validieren
'   Drive, Path, Filename, Extension
'Dazu muss erstmal der String in ein parsbares format gebracht werden.

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
    aStrPath = RecursiveReplaceSL(aStrPath, "\\", "\", 2) 'muss bei 2 starten, weil am Anfang ein doppelter "\\" vorkommen darf
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
                            Optional WinApiErr, _
                            Optional bLoud As Boolean = True, _
                            Optional bErrLog As Boolean = True, _
                            Optional vbDecor As VbMsgBoxStyle = vbOKCancel, _
                            Optional bRetry As Boolean) As VbMsgBoxResult
    'ErrHandler = GlobalErrHandler(Me, PrcName, AddInfo, BolLoud, bErrLog, vbDecor)
    If bRetry Then
        
        ErrHandler = MessErrorRetry(TypeName(Me), PrcName, AddInfo, WinApiErr, bErrLog)
        
    Else
        
        ErrHandler = MessError(TypeName(Me), PrcName, AddInfo, WinApiErr, bLoud, bErrLog, vbDecor)
        
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
