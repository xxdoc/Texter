VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsAppData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private iData() As String
Private iName() As String
Private iCount As Long
Private AppDataLocation As String
Private ErrorTime As Date
Private LastError As Long

Public Enum adMessageType
    adError
    adUserNotification
End Enum

Const MaxFileSize = 100000  'in bytes
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, pidl As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Public Property Get AppData(DataName As String) As String
    Dim X As Long
    For X = 1 To iCount
        If LCase(DataName) = LCase(iName(X)) Then
            AppData = iData(X)
            Exit For
        End If
    Next X
End Property

Public Property Let AppData(DataName As String, NewVal As String)
    Dim X As Long
    Dim Found As Boolean
    Dim Changed As Boolean
    For X = 1 To iCount
        If LCase(DataName) = LCase(iName(X)) Then
            If LCase(iData(X)) <> LCase(NewVal) Then
                Changed = True
                iData(X) = NewVal
            End If
            Found = True
            Exit For
        End If
    Next X
    If Not Found Then
        'add line
        iCount = iCount + 1
        ReDim Preserve iName(iCount)
        ReDim Preserve iData(iCount)
        iName(iCount) = DataName
        iData(iCount) = NewVal
        Changed = True
    End If
    If Changed Then SaveAppData
End Property

Private Sub CheckLogSize()
'---------------------------------------------------------------------------------------
' Procedure : CheckLogSize
' Author    : David
' Date      : 2/7/2012
' Purpose   : check if log size is too big and clip if necessary
'---------------------------------------------------------------------------------------
'
    Dim LN() As String
    Dim L As Long
    Dim ST As Long
    Dim NewL As Long
    On Error GoTo ErrExit
    If FileLen(AppDataLocation & "\Log.txt") > MaxFileSize Then
        Open AppDataLocation & "\Log.txt" For Input As #1
        Do While Not EOF(1)
            L = L + 1
            ReDim Preserve LN(L)
            Line Input #1, LN(L)
        Loop
        Close #1
        Open AppDataLocation & "\LogNew.txt" For Output As #1
        'start 10% from beginning and write new file
        ST = L * 0.1
        For NewL = ST To L
            Print #1, LN(NewL)
        Next NewL
        Close #1
        Kill AppDataLocation & "\Log.txt"
        Name AppDataLocation & "\LogNew.txt" As AppDataLocation & "\Log.txt"
    End If
ErrExit:
    Exit Sub
End Sub

Private Sub CheckNewFile()
'---------------------------------------------------------------------------------------
' Procedure : NewFile
' Author    : David
' Date      : 12/23/2011
' Purpose   : initilize a new file
'---------------------------------------------------------------------------------------
    If Not DirExists(AppDataLocation) Then
        MkDirs AppDataLocation
    End If
    If Dir(AppDataLocation & "\AppData.txt") = "" Then
        AppData("DateStart") = Date & " " & Time
    End If
End Sub

Private Sub Class_Initialize()
    Dim P As String
    P = CommonDocs
    If P = "" Then P = App.Path
    AppDataLocation = P & "\" & App.EXEName & "\Common"
    CheckNewFile
    LoadData
End Sub

Public Property Get CommonDocs() As String
    Dim Rtn As Long
    Dim Buffer As String
    Dim ID As Long
    Buffer = Space$(260)
    Rtn = SHGetSpecialFolderLocation(0&, &H2E, ID)
    If Rtn = 0 Then
        If SHGetPathFromIDList(ID, Buffer) Then
          CommonDocs = TrimNull(Buffer)
        End If
    End If
End Property

Public Property Get DataDir() As String
    DataDir = AppDataLocation
End Property

Private Function DirExists(Path As String) As Boolean
'---------------------------------------------------------------------------------------
' Procedure : DirExists
' Author    : David
' Date      : 12/23/2011
' Purpose   : check if a folder exists by trying to open a file in the Path
'---------------------------------------------------------------------------------------
    Dim TestFile As String
    On Error GoTo ErrHandler
    TestFile = Path & "\Test.txt"
    'check if there already is a test file
    If Dir(TestFile) = "" Then
        Open TestFile For Output As #1
        Close #1
        Kill TestFile
    End If
    DirExists = True
ErrExit:
    Exit Function
ErrHandler:
    DirExists = False
    Resume ErrExit
End Function

Public Sub DisplayError(ByVal ErrorNum As Long, Module As String, Procedure As _
    String, Optional Description As String = "")
    Dim Skip As Boolean
    On Error GoTo ErrExit
    SaveToLog Description, Module, Procedure, ErrorNum, adError
    If ErrorTime = 0 Then ErrorTime = Now
    'check if mutiple errors of the same type should be skipped
    If ErrorNum <> 0 Then
        If ErrorNum = LastError Then
            If Abs(DateDiff("s", ErrorTime, Now)) < 4 Then Skip = True
        End If
    End If
    If Not Skip Then
        MsgBox "Error Number:" & vbTab & ErrorNum & Chr$(13) _
                & "Module:     " & vbTab & Module & Chr$(13) _
                & "Procedure:" & vbTab & Procedure & Chr$(13) _
                & "Description:" & vbTab & Description & Chr$(13)
        ErrorTime = Now
        LastError = ErrorNum
    End If
ErrExit:
    Exit Sub
End Sub

Private Sub GetFormData(Frm As Form, LeftPos As Long, TopPos As Long, State As Long, frmWidth As Long, frmHeight As Long)
    Dim DN As String
    DN = "Form_" & Frm.Name
    LeftPos = Val(AppData(DN & "_Left"))
    TopPos = Val(AppData(DN & "_Top"))
    State = Val(AppData(DN & "_State"))
    frmWidth = Val(AppData(DN & "_Width"))
    frmHeight = Val(AppData(DN & "_Height"))
End Sub

Private Sub LoadData()
'---------------------------------------------------------------------------------------
' Procedure : LoadData
' Author    : David
' Date      : 12/4/2010
' Purpose   :
'---------------------------------------------------------------------------------------
    iCount = 0
    Open AppDataLocation & "\AppData.txt" For Input As #1
    Do While Not EOF(1)
        iCount = iCount + 1
        ReDim Preserve iData(iCount)
        ReDim Preserve iName(iCount)
        Input #1, iName(iCount), iData(iCount)
    Loop
    Close #1
End Sub

Public Sub LoadFormData(Frm As Form, Optional LoadSizeData As Boolean = False)
    Dim LP As Long
    Dim TP As Long
    Dim ST As Long
    Dim Wd As Long
    Dim Ht As Long
    GetFormData Frm, LP, TP, ST, Wd, Ht
    'don't resize fixed border windows
    If Frm.BorderStyle <> 3 And ST = vbNormal And LoadSizeData Then
        If Wd > 0 Then Frm.Width = Wd
        If Ht > 0 Then Frm.Height = Ht
    End If
    If (LP + Frm.Width > Screen.Width) Or LP = 0 Then
        LP = (Screen.Width - Frm.Width) / 2
        If LP < 0 Then LP = 0
    End If
    If (TP + Frm.Height > Screen.Height) Or TP = 0 Then
        TP = (Screen.Height - Frm.Height) / 2
        If TP < 0 Then TP = 0
    End If
    Frm.Left = LP
    Frm.Top = TP
    Frm.WindowState = ST
End Sub

Private Function MkDirs(ByVal Folder As String) As Boolean
   Dim F() As String
   Dim attr As Long
   Dim first As Long
   Dim I As Long

   ' Split incoming folder into subfolders.
   F = Split(Folder, "\")
   For I = 1 To UBound(F)
      F(I) = F(I - 1) & "\" & F(I)
   Next I

   ' If the input path is UNC, the first element
   ' will be empty and the second "\", so we need
   ' to adjust where we start creating folders.
   If F(0) = "" And UBound(F) > 0 Then
      If F(1) = "\" Then
         first = 4  'fourth element is first path.
      End If
   End If

   ' Use errors to signal need to take action.
   On Error Resume Next
   For I = first To UBound(F)
      ' Check if this level already exists.
      attr = GetAttr(F(I))
      If Err.Number Then
         ' Folder likely doesn't exist,
         ' clear error and create.
         Err.Clear
         MkDir F(I)
         If Err.Number Then Exit For
      End If
   Next I

   ' Return success?
   MkDirs = CBool(GetAttr(Folder) And vbDirectory)
End Function

Private Sub SaveAppData()
'---------------------------------------------------------------------------------------
' Procedure : SaveAppData
' Author    : David
' Date      : 12/4/2010
' Purpose   :
'---------------------------------------------------------------------------------------
    Dim C As Long
    Dim Nm As String
    Nm = AppDataLocation & "\AppData.txt"
    Open AppDataLocation & "\AppDataNew.txt" For Output As #1
    For C = 1 To iCount
        If iData(C) <> "" Then
            Write #1, iName(C), iData(C)
        End If
    Next C
    Close #1
    On Error Resume Next
    Kill Nm
    On Error GoTo 0
    Name AppDataLocation & "\AppDataNew.txt" As Nm
End Sub

Public Sub SaveFormData(Frm As Form)
    Dim DN As String
    DN = "Form_" & Frm.Name
    If Frm.WindowState <> vbMinimized Then
        AppData(DN & "_Left") = Frm.Left
        AppData(DN & "_Top") = Frm.Top
        AppData(DN & "_State") = Frm.WindowState
        AppData(DN & "_Width") = Frm.Width
        AppData(DN & "_Height") = Frm.Height
    End If
End Sub

Public Sub SaveToLog(Message As String, Optional Module As String, _
    Optional Procedure As String, Optional ErrorNum As Long, _
    Optional MessageType As adMessageType = adUserNotification)
    On Error GoTo ErrExit:
    If adError Then
        Open AppDataLocation & "\Log.txt" For Append As #1
        Write #1, " ----------- Error -----------"
        Write #1, Format(Now, " dd-MMM-yy h:mm:ss AM/PM")
        Write #1, " Description: " & Message
        Write #1, " Module: " & Module
        Write #1, " Procedure: " & Procedure
        Write #1, " Error #: " & ErrorNum
        Write #1, " -----------------------------"
        Close #1
    Else
        Open AppDataLocation & "\Log.txt" For Append As #1
        Write #1, Message
        Close #1
    End If
    CheckLogSize
ErrExit:
    Exit Sub
End Sub

Public Function TrimNull(ByVal StrZ As String) As String
   TrimNull = Left$(StrZ, InStr(StrZ & vbNullChar, vbNullChar) - 1)
End Function

