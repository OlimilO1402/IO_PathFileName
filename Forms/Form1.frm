VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Testing class PathFileName"
   ClientHeight    =   9780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10335
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9780
   ScaleWidth      =   10335
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton BtnOpenExplorer 
      Caption         =   "Open Explorer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   43
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton BtnPathJoin 
      Caption         =   "Path Join ..\"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   47
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton BtnFileAttributes 
      Caption         =   "File Attributes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   46
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton BtnTestUnicode 
      Caption         =   "TestUnicode >>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   45
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton BtnInfo 
      Caption         =   "Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   42
      Top             =   6840
      Width           =   1815
   End
   Begin VB.CommandButton BtnPathFileNameDelete 
      Caption         =   "Delete File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   40
      Top             =   6360
      Width           =   1815
   End
   Begin VB.CommandButton BtnPathFileNameMove 
      Caption         =   "Move File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   39
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Copy File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   44
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton BtnPathFileNameCreate 
      Caption         =   "Create File"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   38
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton BtnPathDelete 
      Caption         =   "Delete Path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   41
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton BtnPathCreate 
      Caption         =   "Create Path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   36
      Top             =   3360
      Width           =   1815
   End
   Begin VB.CommandButton BtnTestStartWaitCalc 
      Caption         =   "Start ClickMe && Wait"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   33
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton BtnTestStartCalc 
      Caption         =   "Start Calc.exe"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   32
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton BtnTestExists 
      Caption         =   "File Exists?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   34
      Top             =   1680
      Width           =   1815
   End
   Begin VB.CommandButton BtnTestPathExists 
      Caption         =   "Path Exists?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   37
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton BtnIsPathOrFile 
      Caption         =   "IsPath/IsFile"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   31
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton BtnUserPath 
      Caption         =   "User-Path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   30
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton BtnTempPath 
      Caption         =   "Temp-Path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8400
      TabIndex        =   29
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton BtnFileName 
      Caption         =   "/"
      Height          =   255
      Left            =   5160
      TabIndex        =   25
      ToolTipText     =   "Edit"
      Top             =   3000
      Width           =   375
   End
   Begin VB.CommandButton BtnExtension 
      Caption         =   "/"
      Height          =   255
      Left            =   5160
      TabIndex        =   24
      ToolTipText     =   "Edit"
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton BtnFileNameOnly 
      Caption         =   "/"
      Height          =   255
      Left            =   5160
      TabIndex        =   23
      ToolTipText     =   "Edit"
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton BtnPath 
      Caption         =   "/"
      Height          =   255
      Left            =   5160
      TabIndex        =   22
      ToolTipText     =   "Edit"
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton BtnPathOnly 
      Caption         =   "/"
      Height          =   255
      Left            =   5160
      TabIndex        =   21
      ToolTipText     =   "Edit"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton BtnPFN 
      Caption         =   "/"
      Height          =   255
      Left            =   7920
      TabIndex        =   20
      ToolTipText     =   "Edit"
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox TxtPFN 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   120
      Width           =   6615
   End
   Begin VB.CommandButton BtnDrive 
      Caption         =   "/"
      Height          =   255
      Left            =   5160
      TabIndex        =   18
      ToolTipText     =   "Edit"
      Top             =   1200
      Width           =   375
   End
   Begin VB.TextBox TxtFileName 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   3000
      Width           =   3855
   End
   Begin VB.TextBox TxtExtension 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2640
      Width           =   3855
   End
   Begin VB.TextBox TxtFileNameOnly 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2280
      Width           =   3855
   End
   Begin VB.TextBox TxtPath 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1920
      Width           =   3855
   End
   Begin VB.TextBox TxtPathOnly 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox TxtDrive 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   1200
      Width           =   3855
   End
   Begin VB.ListBox LstPaths 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   5640
      TabIndex        =   5
      Top             =   1560
      Width           =   2655
   End
   Begin VB.ListBox LstBsps 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6360
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   8175
   End
   Begin VB.Label LblPathOrFile 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   35
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Lbl1Shorted 
      AutoSize        =   -1  'True
      Caption         =   "Shorted:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   28
      Top             =   840
      Width           =   720
   End
   Begin VB.Label Lbl1Quoted 
      AutoSize        =   -1  'True
      Caption         =   "PFN.Quoted:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   27
      Top             =   480
      Width           =   1065
   End
   Begin VB.Label LblPathFileName1 
      AutoSize        =   -1  'True
      Caption         =   "PathFileName:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   1155
   End
   Begin VB.Label LblPFNQuoted 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1320
      TabIndex        =   11
      Top             =   480
      Width           =   525
   End
   Begin VB.Label LblPFNShorted 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1320
      TabIndex        =   10
      Top             =   840
      Width           =   525
   End
   Begin VB.Label LblFileNameOnly1 
      AutoSize        =   -1  'True
      Caption         =   "FileNameOnly:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   1140
   End
   Begin VB.Label LblPathOnly1 
      AutoSize        =   -1  'True
      Caption         =   "PathOnly:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   8
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label LblPathCount 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6600
      TabIndex        =   7
      Top             =   1200
      Width           =   555
   End
   Begin VB.Label LblPathCount1 
      AutoSize        =   -1  'True
      Caption         =   "PathCount:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5640
      TabIndex        =   6
      Top             =   1200
      Width           =   945
   End
   Begin VB.Label LblExtension1 
      AutoSize        =   -1  'True
      Caption         =   "Extension:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label LblFileName1 
      AutoSize        =   -1  'True
      Caption         =   "FileName:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   780
   End
   Begin VB.Label LblPath1 
      AutoSize        =   -1  'True
      Caption         =   "Path:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   435
   End
   Begin VB.Label LblDrive1 
      AutoSize        =   -1  'True
      Caption         =   "Drive:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_PFN As PathFileName

Private Sub BtnInfo_Click()
    MsgBox App.CompanyName & " " & App.EXEName & " v" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & App.FileDescription
End Sub

Private Sub BtnIsPathOrFile_Click()
    MsgBox "it's a " & IIf(m_PFN.IsFile, "file", "path") & vbCrLf & m_PFN.Value
End Sub

Private Sub BtnOpenExplorer_Click()
    If Not m_PFN.PathExists Then MsgBox "Path does not exist: " & vbCrLf & m_PFN.Path: Exit Sub
    Shell "Explorer.exe " & m_PFN.Path, vbNormalFocus
End Sub

Private Sub BtnPathFileNameCreate_Click()
    If m_PFN.Exists Then MsgBox "File does already exist: " & vbCrLf & m_PFN.Value: Exit Sub
     m_PFN.WriteStr "Just a test string"
     m_PFN.CloseFile
End Sub

Private Sub BtnPathFileNameMove_Click()
    If Not m_PFN.Exists Then
        If MsgBox("Nothing to move, file not found, create a file? " & vbCrLf & m_PFN.Value, vbOKCancel) = vbCancel Then Exit Sub
        BtnCreatePathFileName_Click
    End If
    'verschieben wohin?
    Dim Path As String: Path = InputBox("Edit the path where to move the file:", , m_PFN.Path)
    If StrPtr(Path) = 0 Then Exit Sub 'Cancel
    If UCase(m_PFN.Path) <> UCase(Path) Then
        Dim pathToMove As PathFileName: Set pathToMove = MNew.PathFileName(Path)
        If Not pathToMove.Exists Then
            If MsgBox("Path does not exist create it? " & pathToMove.Value, vbOKCancel) = vbCancel Then Exit Sub
            pathToMove.PathCreate
        End If
        Set m_PFN = m_PFN.MoveTo(pathToMove)
    End If
End Sub

Private Sub BtnPathFileNameDelete_Click()
    'Delete only if the file exists
    If MsgBox("Are you sure you want to delete the file: " & vbCrLf & m_PFN.Value) = vbCancel Then Exit Sub
    If m_PFN.Exists Then
        If m_PFN.Delete Then
            MsgBox "OK file successfully deleted: " & vbCrLf & m_PFN.Value
        Else
            MsgBox "Could not delete the file: " & vbCrLf & m_PFN.Value
        End If
    Else
        MsgBox "File not found, or file does not exist, nothing to delete: " & vbCrLf & m_PFN.Value
    End If
End Sub

Private Sub BtnPathJoin_Click()
    
    Dim basepath As PathFileName: Set basepath = MNew.PathFileName(App.Path)
    Dim spfncls  As String:            spfncls = "..\..\MyRepo\Classes\MyClass1.cls"
    Dim pfn_cls  As PathFileName:  Set pfn_cls = MNew.PathFileName(spfncls)
    
    pfn_cls.PathJoin basepath
    
    MsgBox basepath.Value & " & " & spfncls & " = " & vbCrLf & _
            pfn_cls.Value
        
End Sub

Private Sub BtnTestPathExists_Click()
    MsgBox "PathExists? " & vbCrLf & m_PFN.Path & vbCrLf & m_PFN.PathExists
End Sub

Private Sub BtnTestExists_Click()
    MsgBox "Exists? " & vbCrLf & m_PFN.Value & vbCrLf & m_PFN.Exists
End Sub

Private Sub BtnTestStartCalc_Click()
    Dim CalcExe As PathFileName: Set CalcExe = MNew.PathFileName("Calc.exe")
    CalcExe.Start
End Sub

Private Sub BtnTestStartWaitCalc_Click()
    Dim ClickMe As PathFileName: Set ClickMe = MNew.PathFileName(App.Path & "\ClickMe\ClickMe.exe")
    'Dim ClickMe As PathFileName: Set ClickMe = MNew.PathFileName("C:\TestDir\ClickMe\ClickMe.exe")
    If Not ClickMe.Exists Then Set ClickMe = MNew.PathFileName(App.Path & "\ClickMe.exe")
    If Not ClickMe.Exists Then MsgBox "ClickMe.exe not found: " & vbCrLf & ClickMe.Value: Exit Sub
    ClickMe.StartWait
    MsgBox "Program ClickMe.exe terminated", , Me.Caption & " function StartWait"
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

Private Sub BtnTestUnicode_Click()
    Form2.Show vbModal, Me
End Sub

Private Sub BtnFileAttributes_Click()
    If m_PFN.Exists Then
        MsgBox m_PFN.AttributesToStr
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision
    Call AddExamples:
    CreatePFN 0
End Sub
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
    'complete path: file on a local drive
    LstBsps.AddItem "C:\Hauptverzeichnis\Unterverzeichnis\Datei.txt"
    LstBsps.AddItem "\\SOLS_DS\Daten\TestDir\Unterverzeichnis\Datei.txt"
    'complete path: file on a network-drive
    LstBsps.AddItem "\\Server\Hauptverzeichnis\Unterverzeichnis\Datei.txt"
    
    'if parts are missing the path is relative
    LstBsps.AddItem "SubFolder\File.txt"
    LstBsps.AddItem "..\SubFolder\File.txt"
    
    
    'special parsing obstacles
    LstBsps.AddItem "    \\ C:\ Dingspfad?=*+'# \\ DingsDAtei-.,;:_#+'*~´ß0987654321^°!""§$%&/()=?`\}][{³²"
    
    'only drive
    LstBsps.AddItem "C:"
    LstBsps.AddItem "C:\"
    LstBsps.AddItem "\\Server"
    LstBsps.AddItem "\\Server\"
    
    'drive is missing
    LstBsps.AddItem "Hauptverzeichnis\Unterverzeichnis\Datei.txt"
    LstBsps.AddItem "\Hauptverzeichnis\Unterverzeichnis\Datei.txt"

    'only path
    LstBsps.AddItem "Hauptverzeichnis\"
    LstBsps.AddItem "\Hauptverzeichnis\"
    LstBsps.AddItem "Hauptverzeichnis\Unterverzeichnis\"
    LstBsps.AddItem "\Hauptverzeichnis\Unterverzeichnis\"
    
    'path is missing
    LstBsps.AddItem "C:\Datei.txt"
    LstBsps.AddItem "\\Server\Datei.txt"
    
    'only file
    LstBsps.AddItem "Datei"
    LstBsps.AddItem "\Datei"
    
    'file is missing
    LstBsps.AddItem "C:\Hauptverzeichnis\.txt"
    LstBsps.AddItem "\\Server\Hauptverzeichnis\.txt"
    LstBsps.AddItem "C:\Hauptverzeichnis\Unterverzeichnis\.txt"
    LstBsps.AddItem "\\Server\Hauptverzeichnis\Unterverzeichnis\.txt"
    
    'only extension
    LstBsps.AddItem ".txt"
    LstBsps.AddItem ".Testtxt"
    
    'extension is missing
    LstBsps.AddItem "C:\Hauptverzeichnis\Datei"
    LstBsps.AddItem "\\Server\Hauptverzeichnis\Datei"
    LstBsps.AddItem "C:\Hauptverzeichnis\Unterverzeichnis\Datei"
    LstBsps.AddItem "\\Server\Hauptverzeichnis\Unterverzeichnis\Datei"
    
    'drive and path are missing
    LstBsps.AddItem "Datei.txt"
    LstBsps.AddItem "Datei.Testtxt"
    LstBsps.AddItem "\Datei.txt"
    LstBsps.AddItem "\Datei.Testtxt"
    
    'drive and file are missing
    LstBsps.AddItem "Hauptverzeichnis\.txt"
    LstBsps.AddItem "Hauptverzeichnis\.Testtxt"
    LstBsps.AddItem "Hauptverzeichnis\Unterverzeichnis\.txt"
    LstBsps.AddItem "Hauptverzeichnis\Unterverzeichnis\.Testtxt"
    LstBsps.AddItem "\Hauptverzeichnis\Unterverzeichnis\.txt"
    LstBsps.AddItem "\Hauptverzeichnis\Unterverzeichnis\.Testtxt"
    
    'drive and extension are missing
    LstBsps.AddItem "Hauptverzeichnis\Datei"
    LstBsps.AddItem "\Unterverzeichnis\Datei"
    LstBsps.AddItem "Hauptverzeichnis\Unterverzeichnis\Datei"
    LstBsps.AddItem "\Hauptverzeichnis\Unterverzeichnis\Datei"
    
    'path and file are missing
    LstBsps.AddItem "C:\.txt"
    LstBsps.AddItem "\\Server\.txt"
    LstBsps.AddItem "C:\.Testtxt"
    LstBsps.AddItem "\\Server\.Testtxt"

    'path and extension are missing
    LstBsps.AddItem "C:\Datei"
    LstBsps.AddItem "\\Server\Datei"

    'file and extension are missing
    LstBsps.AddItem "C:\Hauptverzeichnis\"
    LstBsps.AddItem "\\Server\Hauptverzeichnis\"
    LstBsps.AddItem "C:\Hauptverzeichnis\Unterverzeichnis\"
    LstBsps.AddItem "\\Server\Hauptverzeichnis\Unterverzeichnis\"
    
    'one directory higher
    LstBsps.AddItem ".\Unterverzeichnis\Datei"
    LstBsps.AddItem ".\Unterverzeichnis\Datei."
    LstBsps.AddItem ".\Unterverzeichnis\Datei.txt"
    LstBsps.AddItem ".\Hauptverzeichnis\Unterverzeichnis\Datei"
    LstBsps.AddItem ".\Hauptverzeichnis\Unterverzeichnis\Datei."
    LstBsps.AddItem ".\Hauptverzeichnis\Unterverzeichnis\Datei.txt"
    
    'one directory higher
    LstBsps.AddItem "..\Unterverzeichnis\Datei."
    LstBsps.AddItem "..\Unterverzeichnis\Datei.txt"
    LstBsps.AddItem "..\Hauptverzeichnis\Unterverzeichnis\Datei"
    LstBsps.AddItem "..\Hauptverzeichnis\Unterverzeichnis\Datei."
    LstBsps.AddItem "..\Hauptverzeichnis\Unterverzeichnis\Datei.txt"
    
    'for stressing reasons we add a little bit of brainfuck ;-)
    LstBsps.AddItem ".:\.:.:.:..\:\.\\:\\.....\\.\..\\\.\"
    LstBsps.AddItem ".:\.:.:.:..\:\.\\:\\.....\\.\..\\\.\"
    LstBsps.AddItem "A:\.:.:.:..\:\.\\:\\.....\\.\..\\\.\"
    LstBsps.AddItem "\\.:.:.:..\:\.\\:\\.....\\.\..\\\.\"
        
End Sub
