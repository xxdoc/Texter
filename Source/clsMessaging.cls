VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMessaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private modSendSaved As Boolean
Private modSendMessage As String
Private modSendFileName As String
Private modSendFolderLocation As String
Private modReadFileName As String
Const CSIDL_COMMON_DOCUMENTS = &H2E
Private Declare Function SHGetSpecialFolderLocation Lib "shell32" (ByVal hWndOwner As Long, ByVal nFolder As Long, PIDL As Long) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal PIDL As Long, ByVal pszPath As String) As Long

Private Sub Class_Initialize()
    SendFolderLocation = CommonDocs
End Sub

Private Property Get CommonDocs() As String
    Dim Rtn As Long
    Dim Buffer As String
    Dim ID As Long
    Buffer = Space$(260)
    Rtn = SHGetSpecialFolderLocation(0&, CSIDL_COMMON_DOCUMENTS, ID)
    If Rtn = 0 Then
        If SHGetPathFromIDList(ID, Buffer) Then
          CommonDocs = TrimNull(Buffer)
        End If
    End If
End Property

Private Sub EraseOldSend()
    Dim FN As Integer
    If Dir(SendFileFullName) <> "" Then Kill SendFileFullName
    FN = FreeFile
    Open SendFileFullName For Output As #FN
    Close #FN
    SetAttr SendFileFullName, vbNormal
End Sub

Public Function FileChanged() As Boolean
    On Error GoTo ErrExit
    'check for archive bit set
    FileChanged = ((GetAttr(modReadFileName) And vbArchive) <> 0)
    On Error GoTo 0
ErrExit:
    Exit Function
End Function

Private Function IsTextFile(NewFile As String) As Boolean
    On Error GoTo ErrExit
    IsTextFile = (LCase(Right(NewFile, 4)) = ".txt")
ErrExit:
    Exit Function
End Function

Public Property Get ReadFileName() As String
    ReadFileName = modReadFileName
End Property

Public Property Let ReadFileName(NewVal As String)
    If IsTextFile(NewVal) Then
        modReadFileName = NewVal
        ReadMessage 'to clear old data
    Else
        Err.Raise 52, , "Not a text file. Must end in '.txt'"
    End If
End Property

Public Function ReadMessage() As String
    Dim FN As Integer
    Dim FileLen As Long
    On Error GoTo ErrExit
    ReadMessage = ""
    FN = FreeFile
    Open modReadFileName For Input Lock Read Write As FN
    FileLen = LOF(FN)
    If FileLen > 0& Then ReadMessage = Input(FileLen, FN)
    Close #FN
    'clear archive bit
    SetAttr modReadFileName, vbNormal
    On Error GoTo 0
ErrExit:
End Function

Private Function SaveSendToFile() As Boolean
    Dim FN As Integer
    On Error GoTo ErrExit
    SaveSendToFile = False
    FN = FreeFile
    If (GetAttr(SendFileFullName) And vbArchive) = 0 Then
        'file has been read, overwrite
        Open SendFileFullName For Output Lock Read Write As #FN
    Else
        'file has not been read, append
        Open SendFileFullName For Append Lock Read Write As #FN
    End If
    Print #FN, modSendMessage & vbNewLine
    Close #FN
    SaveSendToFile = True
ErrExit:
End Function

Public Function SendFileFullName() As String
    SendFileFullName = modSendFolderLocation & "\" & modSendFileName
End Function

Public Property Get SendFilename() As String
    SendFilename = modSendFileName
End Property

Public Property Let SendFilename(NewVal As String)
    If IsTextFile(NewVal) Then
        modSendFileName = NewVal
        EraseOldSend
    Else
        Err.Raise 52, , "Not a text file. Must end in '.txt'"
    End If
End Property

Public Property Get SendFolderLocation() As String
    SendFolderLocation = modSendFolderLocation
End Property

Public Property Let SendFolderLocation(NewVal As String)
    modSendFolderLocation = NewVal
End Property

Public Function SendMessage(Message As String) As Boolean
    If modSendSaved Then
        'save new message
        modSendMessage = Message
    Else
        'add new message to old message and save
        modSendMessage = modSendMessage & Message
    End If
    If modSendMessage <> "" Then
        modSendSaved = SaveSendToFile
        SendMessage = modSendSaved
    End If
End Function

Private Function TrimNull(ByVal StrZ As String) As String
   TrimNull = Left$(StrZ, InStr(StrZ & vbNullChar, vbNullChar) - 1)
End Function

