VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Texter"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6780
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2520
      Top             =   120
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   7050
      Width           =   6780
      _ExtentX        =   11959
      _ExtentY        =   661
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton butSetup 
      Caption         =   "Setup"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   6555
      Width           =   975
   End
   Begin VB.CommandButton butStream 
      Caption         =   "Stream File"
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      ToolTipText     =   "As file changes send text updates."
      Top             =   6555
      Width           =   1575
   End
   Begin VB.TextBox tbEvents 
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   4080
      Width           =   6495
   End
   Begin VB.CommandButton butFile 
      Caption         =   "Choose File"
      Height          =   375
      Left            =   1140
      TabIndex        =   7
      Top             =   6555
      Width           =   1335
   End
   Begin VB.TextBox Textbox 
      Alignment       =   2  'Center
      Height          =   790
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2520
      Width           =   6480
   End
   Begin VB.TextBox tbMessage 
      Height          =   1215
      Left            =   120
      MaxLength       =   160
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   6495
   End
   Begin VB.CommandButton butSingle 
      Caption         =   "Send Message"
      Height          =   375
      Left            =   4140
      TabIndex        =   9
      Top             =   6555
      Width           =   1335
   End
   Begin VB.CommandButton butClose 
      Caption         =   "Exit"
      Height          =   375
      Left            =   5520
      TabIndex        =   10
      Top             =   6555
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   3840
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label13 
      Caption         =   "Event Log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Message"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label11 
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   2
      Top             =   2167
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Stream As Boolean
Dim WithEvents TR As clsTexting
Attribute TR.VB_VarHelpID = -1
Dim StreamCount As Long
Dim LastModified As Date
Dim ErrorCount As Long

Private Sub butClose_Click()
    Unload Me
End Sub

Private Sub butFile_Click()
    On Error GoTo ErrHandler
    Dim F As String
    F = "Text files (.txt)|*.txt"
    With Dialog
        .FileName = AD.AppData("TextFile")
        .InitDir = AD.DataDir
        .Filter = F
        .Flags = cdlOFNHideReadOnly
        .CancelError = True
        .DialogTitle = "Open"
        .ShowOpen
        If CheckFile(.FileName) Then
            AD.AppData("TextFile") = .FileName
            Textbox(4) = .FileName
        Else
            StatusBar1.SimpleText = "Invalid file name."
            Beep
        End If
    End With
ErrExit:
    Exit Sub
ErrHandler:
    Select Case Err.Number
        Case 32755
            'user canceled
        Case Else
            AD.DisplayError Err.Number, "frmMain", "butFile_Click", Err.Description
    End Select
    Resume ErrExit
End Sub

Private Sub butFile_LostFocus()
    StatusBar1.SimpleText = ""
End Sub

Private Sub butSetup_Click()
    frmSetup.Show vbModal
End Sub

Private Sub butSingle_Click()
    Dim C As Long
    On Error GoTo ErrHandler
    SetStreamState False
    StatusBar1.SimpleText = ""
    With TR
        .Message = tbMessage
        .Port = Val(AD.AppData("Port"))
        .Server = AD.AppData("Server")
        .Username = AD.AppData("Username")
        .Password = AD.AppData("Password")
        .FromEmail = AD.AppData("From")
        For C = 0 To 3
            If Val(AD.AppData("Select" & C)) = 1 Then
                .CellNumber = AD.AppData("Cell" & C)
                .Provider = AD.AppData("Provider" & C)
                .Send
                DoEvents
            End If
        Next C
    End With
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    Select Case Err.Number
        Case 380
            StatusBar1.SimpleText = Err.Description
            Beep
        Case Else
            AD.DisplayError Err.Number, "frmMain", "butSingle_Click", Err.Description
    End Select
    Resume ErrExit
End Sub

Private Sub butSingle_LostFocus()
    StatusBar1.SimpleText = ""
End Sub

Private Sub butStream_Click()
    On Error GoTo ErrHandler
    SetStreamState Not Stream
    StatusBar1.SimpleText = ""
    StreamCount = 0
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmMain", "butStream_Click", Err.Description
End Sub

Private Sub butStream_LostFocus()
    StatusBar1.SimpleText = ""
End Sub

Private Function CheckFile(NewFile As String) As Boolean
    On Error GoTo ErrHandler
    CheckFile = False
    If IsTextFile(NewFile) Then
        If FileExists(NewFile) Then
            CheckFile = True
        Else
            CheckFile = CreateFile(NewFile)
        End If
    End If
    On Error GoTo 0
ErrExit:
    Exit Function
ErrHandler:
     AD.DisplayError Err.Number, "frmMain", "CheckFile", Err.Description
     Resume ErrExit
End Function

Private Function CreateFile(NewFile As String) As Boolean
    On Error GoTo ErrHandler
    Open NewFile For Output As #1
    Close #1
    CreateFile = True
ErrExit:
    Exit Function
ErrHandler:
    CreateFile = False
    Resume ErrExit
End Function

Public Sub EventMessage(ByVal Mes As String, Optional ShowTime As Boolean = True)
'---------------------------------------------------------------------------------------
' Procedure : EventMessage
' Author    : David
' Date      : 3/12/2012
' Purpose   : send status message to the user
'---------------------------------------------------------------------------------------
'
    Dim L As Long
    On Error GoTo ErrHandler
    L = Len(tbEvents)
    If L > 5000 Then
        tbEvents.Text = Right$(tbEvents.Text, 1000)
    End If
    If ShowTime Then
        Mes = Format(Now, "hh:mm:ss  AM/PM") & "    " & Mes
    End If
    tbEvents.Text = tbEvents.Text & Mes & vbNewLine
    tbEvents.SelStart = Len(tbEvents.Text)
    AD.SaveToLog Mes
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmMain", "EventMessage", Err.Description
    Resume ErrExit
End Sub

Private Function FileChanged() As Boolean
    Dim tmpDate As Date
    Dim FileName As String
    On Error GoTo ErrExit
    FileName = AD.AppData("TextFile")
    'first check for archive bit set
    If (GetAttr(FileName) And vbArchive) <> 0 Then
        'check date/time as well
        tmpDate = FileDateTime(FileName)
        If LastModified < tmpDate Then
            'file changed
            LastModified = tmpDate
            FileChanged = True
        End If
    End If
    On Error GoTo 0
ErrExit:
    Exit Function
End Function

Private Function FileExists(NewFile As String) As Boolean
    On Error GoTo ErrHandler
    If Dir(NewFile) <> "" Then FileExists = True
ErrExit:
    Exit Function
ErrHandler:
    FileExists = False
    Resume ErrExit
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim AC As String
    On Error GoTo ErrHandler
    AC = Me.ActiveControl.Name
    'skip multiline textboxes
    If AC <> "tbMessage" Then
        Select Case KeyCode
            Case 38
                'up arrow
                SendKeys ("+{tab}")
            Case 40
                'down arrow
                SendKeys ("{tab}")
        End Select
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmMain", "Form_KeyDown", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrHandler
    Dim AC As String
    On Error GoTo ErrHandler
    AC = Me.ActiveControl.Name
    'skip multiline textboxes
    If AC <> "tbMessage" Then
        If KeyAscii = 13 Then
            'enter
            SendKeys ("{tab}")
            KeyAscii = 0
        End If
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmMain", "Form_KeyPress", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    Me.Caption = "Texter (ver. " & App.Major & "." & App.Minor & "." & App.Revision & ")"
    Set TR = New clsTexting
    Textbox(4) = AD.AppData("TextFile")
    SetLastModified
    AD.LoadFormData Me
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmMain", "Form_Load", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrHandler
    Set TR = Nothing
    AD.SaveFormData Me
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmMain", "Form_Unload", Err.Description
    Resume ErrExit
End Sub

Private Function IsTextFile(NewFile As String) As Boolean
    On Error GoTo ErrExit
    IsTextFile = (LCase(Right(NewFile, 4)) = ".txt")
ErrExit:
    Exit Function
End Function

Private Sub ReadFile()
    Dim TextFile As String
    Dim S As String
    Dim FN As Integer
    Dim FileLen As Long
    Dim C As Long
    On Error GoTo ErrHandler
    TextFile = AD.AppData("TextFile")
    If FileChanged Then
        S = ""
        FN = FreeFile
        Open TextFile For Input Lock Read Write As FN
        FileLen = LOF(FN)
        If FileLen > 0& Then S = Input(FileLen, FN)
        Close #FN
        'clear archive bit
        SetAttr TextFile, vbNormal
        With TR
            .Message = S
            .Port = Val(AD.AppData("Port"))
            .Server = AD.AppData("Server")
            .Username = AD.AppData("Username")
            .Password = AD.AppData("Password")
            .FromEmail = AD.AppData("From")
            For C = 0 To 3
                If Val(AD.AppData("Select" & C)) = 1 Then
                    .CellNumber = AD.AppData("Cell" & C)
                    .Provider = AD.AppData("Provider" & C)
                    .Send
                    DoEvents
                End If
            Next C
        End With
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    Select Case Err.Number
        Case 53
            'file not found, exit
            StatusBar1.SimpleText = "File not found."
            SetStreamState False
            Beep
        Case 70
            'file locked, exit
        Case 380
            'settings error
            StatusBar1.SimpleText = Err.Description
            SetStreamState False
            Beep
        Case Else
            AD.DisplayError Err.Number, "frmMain", "ReadFile", Err.Description
            SetStreamState False
    End Select
    Resume ErrExit
End Sub

Private Sub SetLastModified()
    On Error GoTo ErrExit
    LastModified = FileDateTime(AD.AppData("TextFile"))
ErrExit:
End Sub

Private Sub SetStreamState(NewVal As Boolean)
    On Error GoTo ErrHandler
    Stream = NewVal
    If Stream Then
        butStream.Caption = "Stop Streaming File"
        Timer1.Enabled = True
        EventMessage "Streaming File."
    Else
        butStream.Caption = "Stream File"
        Timer1.Enabled = False
        EventMessage "Streaming stopped."
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmMain", "SetStreamState", Err.Description
     Resume ErrExit
End Sub

Private Sub textbox_GotFocus(Index As Integer)
    On Error GoTo ErrHandler
    Textbox(Index).SelStart = 0
    Textbox(Index).SelLength = Len(Textbox(Index).Text)
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmMain", "textbox_GotFocus", Err.Description
    Resume ErrExit
End Sub

Private Sub Timer1_Timer()
    On Error GoTo ErrHandler
    StreamCount = StreamCount + 1
    StatusBar1.SimpleText = "Stream Loop " & StreamCount
    If StreamCount > 99 Then StreamCount = 0
    ReadFile
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmMain", "Timer1_Timer", Err.Description
     Resume ErrExit
End Sub

Private Sub TR_MessageSent()
    On Error GoTo ErrHandler
    EventMessage "Message sent."
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmMain", "TR_MessageSent", Err.Description
     Resume ErrExit
End Sub

Private Sub TR_SendError(ErrNum As Long, ErrDescription As String)
    Dim TextFile As String
    On Error GoTo ErrHandler
    EventMessage "Send Error. Error # " & ErrNum & ", " & ErrDescription
    ErrorCount = ErrorCount + 1
    If ErrorCount > 4 Then
        SetStreamState False
        ErrorCount = 0
    Else
        If Stream Then
            'reset archive bit for file if streaming
            'this will cause another attempt at streaming the file
            TextFile = AD.AppData("TextFile")
            SetAttr TextFile, vbArchive
            LastModified = 0
        End If
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmMain", "TR_SendError", Err.Description
     Resume ErrExit
End Sub

