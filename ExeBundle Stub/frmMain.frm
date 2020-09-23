VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Exe-Bundle v1.0"
   ClientHeight    =   900
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   900
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'API Declarations
Private Declare Function GetTempPath Lib "kernel32.dll" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function RegisterServiceProcess Lib "kernel32.dll" (ByVal dwProcessId As Long, ByVal dwType As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const STILL_ACTIVE As Long = &H103
Private Const PROCESS_ALL_ACCESS As Long = &H1F0FFF

Private Function StartExec()
    Dim pos As Long, strMess As String, strTitle As String
    Dim strBox As String, intBox As Long
    Dim strBundle As String
    Dim numExe As Integer
    Dim strAppDat As String
    
    On Error GoTo ErrHandler 'Error Handling
    
    'Hide app in Ctrl-Alt-Del list
    'This doesnt work in Win2000,XP and 2003
    'RegisterServiceProcess 0, 1
    
    'Read running file into strAppDat$
    Open cPath & App.EXEName & ".exe" For Binary As #1
        strAppDat$ = Input(LOF(1), #1) 'Read
    Close #1
    
    'Show Message
    'Find [Message] in file
    pos = InStr(1, strAppDat$, "[Message]")
    If pos <> 0 Then 'Only if [Message] exists
        strMess = Mid(strAppDat$, pos + 9, InStr(pos, strAppDat$, "[/Message]") - pos)
        pos = InStr(1, strMess, ",")
        strTitle = Mid(strMess, 1, pos - 1)
        strMess = Mid(strMess, pos + 1) 'Message
        pos = InStr(1, strMess, ",")
        
        strBox = Mid(strMess, 1, pos - 1) 'Type of Message
        Select Case strBox
            Case "vbOkOnly": intBox = vbOKOnly
            Case "vbOKCancel": intBox = vbOKCancel
            Case "vbYesNoCancel": intBox = vbYesNoCancel
            Case "vbYesNo": intBox = vbYesNo
            Case "vbQuestion": intBox = vbQuestion
            Case "vbInformation": intBox = vbInformation
            Case "vbSystemModal": intBox = vbSystemModal
        End Select
        
        strMess = Mid(strMess, pos + 1)
        strMess = Mid(strMess, 1, Len(strMess) - 10)
        
        'Show MsgBox
        ret = MsgBox(strMess, intBox, strTitle)
        
        'If user clicks other than Yes or OK -> End
        If ret <> 1 And ret <> 6 Then 'If Not vbYes or vbOK
            End 'End it
        End If
    End If
    
    
    'Get number of exe's bundled
    pos = InStr(1, strAppDat$, "[Exe-Bundle Info][")
    
    If pos <= 0 Then End 'No EXE's bundled -> End
    
    strBundle$ = Mid(strAppDat$, pos + 18)
    pos = InStr(1, strBundle$, "]")
    numExe = CInt(Mid(strBundle$, 1, pos - 1)) 'Number of EXE's bundled

    'Get exe lengths(sizes)
    Dim exeLen As New Collection 'Collection for EXE Sizes
    Dim i As Integer
    
    strBundle$ = Mid(strBundle$, pos + 1)
    pos = InStr(1, strBundle$, ",")
    For i = 1 To numExe
        If pos <> 0 Then
            exeLen.Add CLng(Mid(strBundle$, 1, pos - 1))
            strBundle$ = Mid(strBundle$, pos + 1)
            pos = InStr(1, strBundle$, ",")
        Else
            exeLen.Add CLng(strBundle$)
        End If
    Next
    
    'Create Exe's and run
    Dim strTemp As String
    
    pos = InStr(1, strAppDat$, "[START]")
    For i = 1 To exeLen.Count
        strBundle$ = getTempFile() 'Get tempfile name
        'Write file
        Open strBundle$ For Binary As #1
            strTemp$ = Mid(strAppDat$, pos + 7, exeLen.Item(i))
            Put #1, , strTemp$
        Close #1
        
        pos = pos + exeLen.Item(i) 'Next file's start position
        
        DoEvents
        
        'Start the file using Shell and wait till it is closed
        pId = Shell(strBundle$, vbNormalFocus)
        prog = OpenProcess(PROCESS_ALL_ACCESS, False, pId)
        GetExitCodeProcess prog, progKill
        Do While progKill = STILL_ACTIVE
            DoEvents
            GetExitCodeProcess prog, progKill
        Loop
        
        'Delete the temp file after it is closed by user
        Do While Dir(strBundle$, vbNormal) <> ""
            DoEvents
            Kill strBundle$
        Loop
    Next
    
    'All EXE's executed and closed by user -> End
    End
    
    'Error Handling
ErrHandler:
    MsgBox "Error Number : " & Err.Number & vbCrLf & _
    Err.Description, , "Exe Bundle v1.0"
    End
End Function

Private Function getTempFile() As String
    'Function to get temporary file
    Dim ret As Long
    Dim lngPath As Long, strPath As String
    lngPath = 256
    strPath = String(256, vbNullChar)
    
    'Get system Temp folder path
    ret = GetTempPath(lngPath, strPath)
    
    If ret <> 0 Then 'Success
        strPath = Left(strPath, ret)
    Else 'If failure, use App.Path
        strPath = cPath
    End If
    
    Randomize
    getTempFile = strPath & CStr(CLng(Rnd * 99999) + 1) & ".exe"
    Do While Dir(getTempFile, vbNormal) <> ""
        Randomize
        getTempFile = strPath & CStr(CLng(Rnd * 99999) + 1) & ".exe"
        DoEvents
    Loop
End Function

Private Function cPath() As String
    'Function to return AppPath
    cPath = App.Path
    If Right(cPath, 1) <> "\" Then cPath = cPath & "\"
End Function

Private Sub Form_Load()
    Call StartExec
End Sub
