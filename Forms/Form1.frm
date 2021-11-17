VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Testing class PathFileName"
   ClientHeight    =   9780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10215
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9780
   ScaleWidth      =   10215
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Delete PathFileName"
      Height          =   375
      Left            =   8280
      TabIndex        =   41
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Move PathFileName"
      Height          =   375
      Left            =   8280
      TabIndex        =   40
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton BtnPathFileNameCreate 
      Caption         =   "Create PathFilename"
      Height          =   375
      Left            =   8280
      TabIndex        =   39
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton BtnPathDelete 
      Caption         =   "Delete Path"
      Height          =   375
      Left            =   8280
      TabIndex        =   42
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton BtnPathCreate 
      Caption         =   "Create Path"
      Height          =   375
      Left            =   8280
      TabIndex        =   37
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton BtnTestStartWaitCalc 
      Caption         =   "Start and wait"
      Height          =   375
      Left            =   8280
      TabIndex        =   33
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton BtnTestStartCalc 
      Caption         =   "Start calc"
      Height          =   375
      Left            =   8280
      TabIndex        =   32
      Top             =   2160
      Width           =   1815
   End
   Begin VB.CommandButton BtnTestExists 
      Caption         =   "Exists"
      Height          =   375
      Left            =   8280
      TabIndex        =   34
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton BtnTestPathExists 
      Caption         =   "Path Exists"
      Height          =   375
      Left            =   8280
      TabIndex        =   38
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton BtnIsPathOrFile 
      Caption         =   "IsPath/IsFile"
      Height          =   375
      Left            =   8280
      TabIndex        =   31
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton BtnUserPath 
      Caption         =   "User-Path"
      Height          =   375
      Left            =   8280
      TabIndex        =   30
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton BtnTempPath 
      Caption         =   "Temp-Path"
      Height          =   375
      Left            =   8280
      TabIndex        =   29
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton BtnFileName 
      Caption         =   "/"
      Height          =   255
      Left            =   5040
      TabIndex        =   25
      ToolTipText     =   "Edit"
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton BtnExtension 
      Caption         =   "/"
      Height          =   255
      Left            =   5040
      TabIndex        =   24
      ToolTipText     =   "Edit"
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton BtnFileNameOnly 
      Caption         =   "/"
      Height          =   255
      Left            =   5040
      TabIndex        =   23
      ToolTipText     =   "Edit"
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton BtnPath 
      Caption         =   "/"
      Height          =   255
      Left            =   5040
      TabIndex        =   22
      ToolTipText     =   "Edit"
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton BtnPathOnly 
      Caption         =   "/"
      Height          =   255
      Left            =   5040
      TabIndex        =   21
      ToolTipText     =   "Edit"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton BtnPFN 
      Caption         =   "/"
      Height          =   255
      Left            =   7800
      TabIndex        =   20
      ToolTipText     =   "Edit"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtPFN 
      Height          =   285
      Left            =   1200
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   120
      Width           =   6615
   End
   Begin VB.CommandButton BtnDrive 
      Caption         =   "/"
      Height          =   255
      Left            =   5040
      TabIndex        =   18
      ToolTipText     =   "Edit"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox TxtFileName 
      Height          =   285
      Left            =   1200
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3000
      Width           =   3855
   End
   Begin VB.TextBox TxtExtension 
      Height          =   285
      Left            =   1200
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2640
      Width           =   3855
   End
   Begin VB.TextBox TxtFileNameOnly 
      Height          =   285
      Left            =   1200
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2280
      Width           =   3855
   End
   Begin VB.TextBox TxtPath 
      Height          =   285
      Left            =   1200
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox TxtPathOnly 
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox TxtDrive 
      Height          =   285
      Left            =   1200
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1200
      Width           =   3855
   End
   Begin VB.ListBox LstPaths 
      Height          =   1815
      Left            =   5520
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
   End
   Begin VB.ListBox LstBsps 
      Height          =   6300
      Left            =   0
      TabIndex        =   0
      Top             =   3480
      Width           =   8175
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   7440
      TabIndex        =   36
      Top             =   840
      Width           =   615
   End
   Begin VB.Label LblPathOrFile 
      Caption         =   "Label1"
      Height          =   255
      Left            =   7440
      TabIndex        =   35
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Lbl1Shorted 
      AutoSize        =   -1  'True
      Caption         =   "Shorted:"
      Height          =   195
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   600
   End
   Begin VB.Label Lbl1Quoted 
      AutoSize        =   -1  'True
      Caption         =   "PFN.Quoted:"
      Height          =   195
      Left            =   120
      TabIndex        =   27
      Top             =   480
      Width           =   930
   End
   Begin VB.Label LblPathFileName1 
      AutoSize        =   -1  'True
      Caption         =   "PathFileName:"
      Height          =   195
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   1035
   End
   Begin VB.Label LblPFNQuoted 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1200
      TabIndex        =   11
      Top             =   480
      Width           =   480
   End
   Begin VB.Label LblPFNShorted 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1200
      TabIndex        =   10
      Top             =   840
      Width           =   480
   End
   Begin VB.Label LblFileNameOnly1 
      AutoSize        =   -1  'True
      Caption         =   "FileNameOnly:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1020
   End
   Begin VB.Label LblPathOnly1 
      AutoSize        =   -1  'True
      Caption         =   "PathOnly:"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   690
   End
   Begin VB.Label LblPathCount 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   6360
      TabIndex        =   7
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label LblPathCount1 
      AutoSize        =   -1  'True
      Caption         =   "PathCount:"
      Height          =   195
      Left            =   5520
      TabIndex        =   6
      Top             =   1200
      Width           =   795
   End
   Begin VB.Label LblExtension1 
      AutoSize        =   -1  'True
      Caption         =   "Extension:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   735
   End
   Begin VB.Label LblFileName1 
      AutoSize        =   -1  'True
      Caption         =   "FileName:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   705
   End
   Begin VB.Label LblPath1 
      AutoSize        =   -1  'True
      Caption         =   "Path:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   375
   End
   Begin VB.Label LblDrive1 
      AutoSize        =   -1  'True
      Caption         =   "Drive:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   420
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_PFN As PathFileName

Private Sub BtnIsPathOrFile_Click()
    MsgBox "it's a " & IIf(m_PFN.IsFile, "file", "path") & vbCrLf & m_PFN.Value
End Sub
Private Sub BtnTestPathExists_Click()
    MsgBox "PathExists? " & vbCrLf & m_PFN.PathExists & vbCrLf & m_PFN.Path
End Sub
Private Sub BtnTestExists_Click()
    MsgBox "Exists? " & vbCrLf & m_PFN.Exists & vbCrLf & m_PFN.Value
End Sub

Private Sub BtnTestStartCalc_Click()
    Dim CalcExe As PathFileName: Set CalcExe = MNew.PathFileName("Calc.exe")
    CalcExe.Start
End Sub

Private Sub BtnTestStartWaitCalc_Click()
    Set m_PFN = MNew.PathFileName(App.Path & "\ClickMe\ClickMe.exe")
    m_PFN.StartWait
    MsgBox "Program ClickMe terminated", , Me.Caption & " function StartWait"
End Sub

Private Sub BtnPathCreate_Click()
    Dim s As String: s = TxtPFN.Text
    If Len(s) = 0 Then Exit Sub
    Set m_PFN = MNew.PathFileName(TxtPFN.Text)
    If m_PFN.PathExists Then
        MsgBox "Path already exists: " & vbCrLf & m_PFN.Path
    Else
        m_PFN.PathCreate
        MsgBox IIf(m_PFN.PathExists, "Path successfully created: ", "Could not create Path: ") & vbCrLf & m_PFN.Path
    End If
End Sub
Private Sub BtnPathDelete_Click()
    Dim s As String: s = TxtPFN.Text
    If Len(s) = 0 Then Exit Sub
    Set m_PFN = MNew.PathFileName(TxtPFN.Text)
    If m_PFN.PathExists Then
        m_PFN.PathDelete
        If m_PFN.PathExists Then
            MsgBox "Could not delete path: " & vbCrLf & m_PFN.Path
        End If
    Else
        MsgBox "There is nothing to delete, there is no path, named: " & vbCrLf & m_PFN.Path
    End If
End Sub

Private Sub BtnCreatePathFileName_Click()
    Dim s As String: s = TxtPFN.Text
    If Len(s) = 0 Then Exit Sub
    Set m_PFN = MNew.PathFileName(TxtPFN.Text)
    If Not m_PFN.PathExists Then m_PFN.PathCreate
    If m_PFN.PathExists Then
        If m_PFN.Exists Then
            MsgBox "PathFileName already exists: " & vbCrLf & m_PFN.Value
        Else
            m_PFN.WriteStr "Testfile"
            MsgBox IIf(m_PFN.Exists, "PathFileName successfully created: ", "Could not create PathFileName: ") & vbCrLf & m_PFN.Value
        End If
    Else
        MsgBox "Could not create Path: " & vbCrLf & m_PFN.Path
    End If
End Sub


Private Sub Form_Load():             Call AddExamples:                          CreatePFN 0: End Sub
Private Sub BtnDrive_Click():        m_PFN.Drive = TxtDrive.Text:               View_Update: End Sub
Private Sub BtnExtension_Click():    m_PFN.Extension = TxtExtension.Text:       View_Update: End Sub
Private Sub BtnFileName_Click():     m_PFN.FileName = TxtFileName.Text:         View_Update: End Sub
Private Sub BtnFileNameOnly_Click(): m_PFN.FileNameOnly = TxtFileNameOnly.Text: View_Update: End Sub
Private Sub BtnPath_Click():         m_PFN.Path = TxtPath.Text:                 View_Update: End Sub
Private Sub BtnPathOnly_Click():     m_PFN.PathOnly = TxtPathOnly.Text:         View_Update: End Sub
Private Sub BtnPFN_Click():      Set m_PFN = MNew.PathFileName(TxtPFN.Text):    View_Update: End Sub
Private Sub BtnTempPath_Click():    MsgBox m_PFN.TempPath: End Sub
Private Sub BtnUserPath_Click():    MsgBox m_PFN.UserPath: End Sub

Private Sub Form_Resize()
    Dim L As Single: L = LstBsps.Left
    Dim T As Single: T = LstBsps.Top
    Dim W As Single: W = LstBsps.Width 'Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight - T
    If W > 0 And H > 0 Then LstBsps.Move L, T, W, H
End Sub

Private Sub LstBsps_Click()
    CreatePFN LstBsps.ListIndex
End Sub
Sub CreatePFN(ByVal i As Long)
    Set m_PFN = MNew.PathFileName(LstBsps.List(i))
    View_Update
End Sub
Private Sub View_Clear()
             TxtPFN.Text = "": LblPFNQuoted.Caption = "": LblPFNShorted.Caption = ""
           TxtDrive.Text = "":     TxtPathOnly.Text = "":          TxtPath.Text = ""
    TxtFileNameOnly.Text = "":    TxtExtension.Text = "":      TxtFileName.Text = ""
    LblPathCount.Caption = "":    LstPaths.Clear
End Sub
Private Sub View_Update()
    Dim li As Long: li = LstPaths.ListIndex
    View_Clear
    TxtPFN.Text = m_PFN.Value
    LblPFNQuoted.Caption = m_PFN.Quoted
    LblPFNShorted.Caption = m_PFN.Shorted(35)
    TxtDrive.Text = m_PFN.Drive
    TxtPathOnly.Text = m_PFN.PathOnly
    TxtPath.Text = m_PFN.Path
    TxtFileNameOnly.Text = m_PFN.FileNameOnly
    TxtExtension.Text = m_PFN.Extension
    TxtFileName.Text = m_PFN.FileName
    LblPathCount.Caption = m_PFN.PathCount
    LblPathOrFile.Caption = IIf(m_PFN.IsFile, "file", "path")
    
    'li = LstBsps.ListIndex
    'LstBsps.List(li) = m_PFN.Value
    
    Dim i As Long
    For i = 0 To m_PFN.PathCount - 1
        LstPaths.AddItem m_PFN.PathI(i)
    Next
    If li < LstPaths.ListCount Then
        LstPaths.ListIndex = li
    End If
End Sub

Private Sub LstPaths_DblClick()
    Dim i As Long: i = LstPaths.ListIndex
    If i < 0 Then Exit Sub
    Dim p As String: p = LstPaths.List(i)
    p = InputBox("Pfad", "Editieren", p)
    If StrPtr(p) = 0 Then Exit Sub
    m_PFN.PathI(i) = p
    View_Update
End Sub

Private Sub AddExamples()
    View_Clear
    'vollständiger Pfad: Datei auf lokalem Laufwerk
    LstBsps.AddItem "C:\Hauptverzeichnis\Unterverzeichnis\Datei.txt"
    LstBsps.AddItem "\\SOLS_DS\Daten\TestDir\Unterverzeichnis\Datei.txt"
    'vollständiger Pfad: Datei auf Netzwerk-Laufwerk
    LstBsps.AddItem "\\Server\Hauptverzeichnis\Unterverzeichnis\Datei.txt"
    'Wenn Teile fehlen
    
    'special parsing obstacles
    LstBsps.AddItem "    \\ C:\ Dingspfad?=*+'# \\ DingsDAtei-.,;:_#+'*~´ß0987654321^°!""§$%&/()=?`\}][{³²"
    
    'nur Drive
    LstBsps.AddItem "C:"
    LstBsps.AddItem "C:\"
    LstBsps.AddItem "\\Server"
    LstBsps.AddItem "\\Server\"
    
    'Drive fehlt
    LstBsps.AddItem "Hauptverzeichnis\Unterverzeichnis\Datei.txt"
    LstBsps.AddItem "\Hauptverzeichnis\Unterverzeichnis\Datei.txt"

    'nur Pfad
    LstBsps.AddItem "Hauptverzeichnis\"
    LstBsps.AddItem "\Hauptverzeichnis\"
    LstBsps.AddItem "Hauptverzeichnis\Unterverzeichnis\"
    LstBsps.AddItem "\Hauptverzeichnis\Unterverzeichnis\"
    
    'Pfad fehlt
    LstBsps.AddItem "C:\Datei.txt"
    LstBsps.AddItem "\\Server\Datei.txt"
    
    'nur Datei
    LstBsps.AddItem "Datei"
    LstBsps.AddItem "\Datei"
    
    'Datei fehlt
    LstBsps.AddItem "C:\Hauptverzeichnis\.txt"
    LstBsps.AddItem "\\Server\Hauptverzeichnis\.txt"
    LstBsps.AddItem "C:\Hauptverzeichnis\Unterverzeichnis\.txt"
    LstBsps.AddItem "\\Server\Hauptverzeichnis\Unterverzeichnis\.txt"
    
    'nur Extension
    LstBsps.AddItem ".txt"
    LstBsps.AddItem ".Testtxt"
    
    'Extension fehlt
    LstBsps.AddItem "C:\Hauptverzeichnis\Datei"
    LstBsps.AddItem "\\Server\Hauptverzeichnis\Datei"
    LstBsps.AddItem "C:\Hauptverzeichnis\Unterverzeichnis\Datei"
    LstBsps.AddItem "\\Server\Hauptverzeichnis\Unterverzeichnis\Datei"
    
    'Drive und Pfad fehlen
    LstBsps.AddItem "Datei.txt"
    LstBsps.AddItem "Datei.Testtxt"
    LstBsps.AddItem "\Datei.txt"
    LstBsps.AddItem "\Datei.Testtxt"
    
    'Drive und Datei fehlen
    LstBsps.AddItem "Hauptverzeichnis\.txt"
    LstBsps.AddItem "Hauptverzeichnis\.Testtxt"
    LstBsps.AddItem "Hauptverzeichnis\Unterverzeichnis\.txt"
    LstBsps.AddItem "Hauptverzeichnis\Unterverzeichnis\.Testtxt"
    LstBsps.AddItem "\Hauptverzeichnis\Unterverzeichnis\.txt"
    LstBsps.AddItem "\Hauptverzeichnis\Unterverzeichnis\.Testtxt"
    
    'Drive und Extension fehlen
    LstBsps.AddItem "Hauptverzeichnis\Datei"
    LstBsps.AddItem "\Unterverzeichnis\Datei"
    LstBsps.AddItem "Hauptverzeichnis\Unterverzeichnis\Datei"
    LstBsps.AddItem "\Hauptverzeichnis\Unterverzeichnis\Datei"
    
    'Pfad und Datei fehlen
    LstBsps.AddItem "C:\.txt"
    LstBsps.AddItem "\\Server\.txt"
    LstBsps.AddItem "C:\.Testtxt"
    LstBsps.AddItem "\\Server\.Testtxt"

    'Pfad und Extension fehlen
    LstBsps.AddItem "C:\Datei"
    LstBsps.AddItem "\\Server\Datei"

    'Datei und Extension fehlen
    LstBsps.AddItem "C:\Hauptverzeichnis\"
    LstBsps.AddItem "\\Server\Hauptverzeichnis\"
    LstBsps.AddItem "C:\Hauptverzeichnis\Unterverzeichnis\"
    LstBsps.AddItem "\\Server\Hauptverzeichnis\Unterverzeichnis\"
    
    'ein Verzeichnis Höher
    LstBsps.AddItem ".\Unterverzeichnis\Datei"
    LstBsps.AddItem ".\Unterverzeichnis\Datei."
    LstBsps.AddItem ".\Unterverzeichnis\Datei.txt"
    LstBsps.AddItem ".\Hauptverzeichnis\Unterverzeichnis\Datei"
    LstBsps.AddItem ".\Hauptverzeichnis\Unterverzeichnis\Datei."
    LstBsps.AddItem ".\Hauptverzeichnis\Unterverzeichnis\Datei.txt"

    'ein Verzeichnis Höher
    LstBsps.AddItem "..\Unterverzeichnis\Datei."
    LstBsps.AddItem "..\Unterverzeichnis\Datei.txt"
    LstBsps.AddItem "..\Hauptverzeichnis\Unterverzeichnis\Datei"
    LstBsps.AddItem "..\Hauptverzeichnis\Unterverzeichnis\Datei."
    LstBsps.AddItem "..\Hauptverzeichnis\Unterverzeichnis\Datei.txt"
    
    'for stressing reasons we add a little bit of brainfuck ;-)
    LstBsps.AddItem ".:\.:.:.:..\:\.\\:\\.....\\.\..\\\.\"
    
    'for stressing reasons we add a little bit of brainfuck ;-)
    LstBsps.AddItem ".:\.:.:.:..\:\.\\:\\.....\\.\..\\\.\"
    LstBsps.AddItem "A:\.:.:.:..\:\.\\:\\.....\\.\..\\\.\"
    LstBsps.AddItem "\\.:.:.:..\:\.\\:\\.....\\.\..\\\.\"
        
End Sub
