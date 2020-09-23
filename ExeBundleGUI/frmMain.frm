VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exe-Bundle v1.0"
   ClientHeight    =   5265
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   5730
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   5730
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Message :"
      Height          =   2895
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   5535
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   600
         TabIndex        =   12
         Top             =   1920
         Width           =   3615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton cmdOpenText 
         Caption         =   "&Open"
         Height          =   375
         Left            =   4320
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdClearMess 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Top             =   2400
         Width           =   1095
      End
      Begin VB.TextBox txtMess 
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   240
         Width           =   5295
      End
      Begin VB.Label Label2 
         Caption         =   "Title :"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Type of box to be displayed :"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   2535
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3840
      Top             =   4800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdBundle 
      Caption         =   "&Bundle"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Exe's to be bundled :"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear List"
         Height          =   375
         Left            =   4320
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&Open"
         Height          =   375
         Left            =   4320
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.ListBox List1 
         Height          =   1230
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Exe-Bundle v1.0
'Copyright (c)2005-2006 Sudesh Katariya

Private Sub cmdBundle_Click()
    Dim strDat As String, strTemp As String, strBundle As String
    Dim i As Integer
    
    On Error Resume Next
    
    If List1.ListCount = 0 Then Exit Sub 'No files
    
    'Show save dialog to save final bundled EXE
    With CD
        .CancelError = True
        .DefaultExt = ".exe"
        .Flags = &H2000 Or &H2 Or &H4 Or &H800
        .Filter = "Executables (*.exe)|*.exe"
        .ShowSave
        If Err.Number <> 0 Then Exit Sub 'Exit, user clicked cancel
    End With
    
    'Final Bundled EXE would be like:
    '1. ExeBundleStub Header
    '2. [Message]--data---[/Message]
    '3. [START]
    '4. All bundled files one after another
    '5. [Exe-Bundle Info][no. of files bundled]
    '6. [lengths of files seperated by comma in same order]
    
    
    '1. Put Stub Header
    If Dir(cPath & "ExeBundleStub.exe") = "" Then
        MsgBox "Exe-Bundle Stub not found.", vbCritical, "Error"
        Exit Sub
    End If
    Open cPath & "ExeBundleStub.exe" For Binary As #1
        strTemp$ = Input(LOF(1), #1)
    Close #1
    Open CD.FileName For Binary As #1
        Put #1, , strTemp$
    Close #1
    
    '2 Put Message
    If txtMess.Text <> "" Then 'If message to be displayed
        strTemp$ = "[Message]" & txtTitle.Text & _
        ",vb" & Combo1.Text & "," & txtMess.Text & "[/Message]"
        Open CD.FileName For Binary As #1
            Seek #1, LOF(1) + 1
            Put #1, , strTemp$
        Close #1
    End If
    
    '3. [START]
    strTemp$ = "[START]"
    Open CD.FileName For Binary As #1
        Seek #1, LOF(1) + 1
        Put #1, , strTemp$
    Close #1
    
    '4. All bundled files
    For i = 0 To List1.ListCount - 1
        Open List1.List(i) For Binary As #1
            strTemp$ = Input(LOF(1), #1)
            strBundle$ = strBundle$ & CStr(LOF(1)) & ","
        Close #1
        Open CD.FileName For Binary As #1
            Seek #1, LOF(1) + 1
            Put #1, , strTemp$
        Close #1
        DoEvents
    Next
    
    '5,6. [Exe-Bundle Info]
    strBundle$ = Mid(strBundle$, 1, Len(strBundle$) - 1)
    strBundle$ = "[Exe-Bundle Info][" & CStr(List1.ListCount) & _
    "]" & strBundle$
    Open CD.FileName For Binary As #1
        Seek #1, LOF(1) + 1
        Put #1, , strBundle$
    Close #1
    
    MsgBox "Executables sucessfully bundled.", vbInformation, "SUCCESS"
End Sub

Private Sub cmdClear_Click()
    List1.Clear 'Clear file list
End Sub

Private Sub cmdClearMess_Click()
    txtMess.Text = "" 'Clear message
End Sub

Private Sub cmdOpen_Click()
    'Select file for bundling
    On Error Resume Next
    With CD
        .CancelError = True
        .DefaultExt = ".exe"
        .Flags = &H1000 Or &H4 Or &H800
        .Filter = "Executables (*.exe)|*.exe"
        .ShowOpen
        If Err.Number <> 0 Then Exit Sub
        If Not CheckDuplicate(CD.FileName) Then List1.AddItem CD.FileName
    End With
End Sub

Private Function CheckDuplicate(ByVal strFileName As String) As Boolean
    'Function of checking duplicate in list
    strFileName = LCase(strFileName)
    CheckDuplicate = False
    For i = 0 To List1.ListCount - 1
        If LCase(List1.List(i)) = strFileName Then
            CheckDuplicate = True
            Exit For
        End If
        DoEvents
    Next
End Function

Private Sub cmdOpenText_Click()
    'Select message text from a text file
    On Error Resume Next
    Dim strMess As String
    With CD
        .CancelError = True
        .DefaultExt = ".txt"
        .Flags = &H1000 Or &H4 Or &H800
        .Filter = "Text Documents (*.txt)|*.txt"
        .ShowOpen
        If Err.Number <> 0 Then Exit Sub
        txtMess.Text = ""
        Open .FileName For Input As #1
            strMess = Input(LOF(1), #1)
        Close #1
        If Len(strMess) > 256 Then
            MsgBox "This message cannot be more than 256 chars long.", vbExclamation, "Error"
            Exit Sub
        End If
        txtMess.Text = strMess
    End With
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then End
    With Combo1
        .AddItem "OKOnly"
        .AddItem "OKCancel"
        .AddItem "YesNoCancel"
        .AddItem "YesNo"
        .AddItem "Question"
        .AddItem "Information"
        .AddItem "SystemModal"
    End With
    Combo1.Text = Combo1.List(0)
End Sub

Private Function cPath() As String
    'Function to return AppPath
    cPath = App.Path
    If Right(cPath, 1) <> "\" Then cPath = cPath & "\"
End Function
