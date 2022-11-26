VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   6975
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11370
   LinkTopic       =   "Form2"
   ScaleHeight     =   6975
   ScaleWidth      =   11370
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   0
      MultiLine       =   -1  'True
      OLEDragMode     =   1  'Automatisch
      OLEDropMode     =   1  'Manuell
      ScrollBars      =   3  'Beides
      TabIndex        =   0
      ToolTipText     =   "Drag'n'drop files here"
      Top             =   0
      Width           =   10575
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
    Dim L As Single
    Dim T As Single
    Dim W As Single: W = Me.ScaleWidth
    Dim H As Single: H = Me.ScaleHeight
    If W > 0 And H > 0 Then Text1.Move L, T, W, H
End Sub

Private Sub mnuFileOpen_Click()
    Dim OFD As New OpenFileDialog
    If OFD.ShowDialog(Me) = vbCancel Then Exit Sub
    Dim f As String: f = OFD.FileName
    MString.MsgBoxW f
    
    'Debug.Print f
    Dim pfn As PathFileName: Set pfn = MNew.PathFileName(f)
Try: On Error GoTo Catch
    Text1.Text = pfn.ReadAllText
    'Dim FNr As Integer: FNr = pfn.OpenFile(FileMode_Binary, FileAccess_Read)
    'ReDim Buffer(0 To pfn.Length)
    'Get FNr, , Buffer
    GoTo Finally
Catch:
    'MsgBox "Mist"
Finally:
    pfn.CloseFile
End Sub

Private Sub Text1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not Data.GetFormat(ClipBoardConstants.vbCFFiles) Then Exit Sub
    Dim file As String: file = Data.Files(1)
    Dim pfn As PathFileName: Set pfn = MNew.PathFileName(file)
    Text1.Text = pfn.ReadAllText
End Sub

