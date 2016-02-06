VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTexting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type MesProps
    Port As Long
    Server As String
    Username As String
    Password As String
    CellNum As String
    Provider As String
    Message As String
    FromEmail As String
End Type

Private Enum Props
    PortProp
    Serverprop
    Usernameprop
    PasswordProp
    CellNumProp
    ProviderProp
    MessageProp
    FromProp
End Enum

Event MessageSent()
Event SendError(ErrNum As Long, ErrDescription As String)

Dim Mes As MesProps
Dim cdoMes As CDO.Message
Dim Valid(7) As Boolean
Const MaxLength = 160   'maximum text length

Public Property Get CellNumber() As String
    CellNumber = Mes.CellNum
End Property

Public Property Let CellNumber(NewVal As String)
    'check validity
    NewVal = FixCell(NewVal)
    If Len(NewVal) <> 10 Then Err.Raise 380, , "Invalid cell number."
    Mes.CellNum = NewVal
    Valid(Props.CellNumProp) = True
End Property

Private Sub Class_Initialize()
    Set cdoMes = New CDO.Message
End Sub

Private Sub Class_Terminate()
    Set cdoMes = Nothing
End Sub

Private Function FixCell(OldVal As String) As String
'---------------------------------------------------------------------------------------
' Procedure : FixCell
' Author    : XPMUser
' Date      : 20/Jan/2015
' Purpose   : remove formating from cell number and return just numbers
'---------------------------------------------------------------------------------------
'
    Dim P As Long
    Dim A As Long
    Dim S As String
    FixCell = ""
    For P = 1 To Len(OldVal)
        S = Mid(OldVal, P, 1)
        A = Asc(S)
        If A > 47 And A < 58 Then FixCell = FixCell & S
    Next P
End Function

Public Property Get FromEmail() As String
    FromEmail = Mes.FromEmail
End Property

Public Property Let FromEmail(NewVal As String)
    If NewVal = "" Or Len(NewVal) > 50 Then Err.Raise 380, , "Invalid From Email length."
    Mes.FromEmail = NewVal
    Valid(Props.FromProp) = True
End Property

Public Property Get Message() As String
    Message = Mes.Message
End Property

Public Property Let Message(NewVal As String)
    If NewVal = "" Then Err.Raise 380, , "Invalid Message length."
    If Len(NewVal) > MaxLength Then NewVal = Right(NewVal, MaxLength)
    Mes.Message = NewVal
    Valid(Props.MessageProp) = True
End Property

Private Function MessageValid() As Boolean
    Dim P As Long
    Dim Tmp As Boolean
    MessageValid = False
    Tmp = True
    For P = 0 To 7
        Tmp = Tmp And Valid(P)
    Next P
    MessageValid = Tmp
End Function

Public Property Get Password() As String
    Password = Mes.Password
End Property

Public Property Let Password(NewVal As String)
    If NewVal = "" Or Len(NewVal) > 50 Then Err.Raise 380, , "Invalid Password."
    Mes.Password = NewVal
    Valid(Props.PasswordProp) = True
End Property

Public Property Get Port() As Long
    Port = Mes.Port
End Property

Public Property Let Port(NewVal As Long)
    If NewVal < 0 Or NewVal > 65535 Then Err.Raise 380, , "Invalid Port #."
    Mes.Port = NewVal
    Valid(Props.PortProp) = True
End Property

Public Property Get Provider() As String
    Provider = Mes.Provider
End Property

Public Property Let Provider(NewVal As String)
    If NewVal = "" Or Len(NewVal) > 50 Then Err.Raise 380, , "Invalid Provider."
    Mes.Provider = NewVal
    Valid(Props.ProviderProp) = True
End Property

Public Sub Send()
    On Error GoTo ErrHandler
    If MessageValid Then
        With cdoMes
            With .Configuration.Fields
                .Item(cdoSendUsingMethod).Value = cdoSendUsingPort
                .Item(cdoSMTPUseSSL).Value = True
                .Item(cdoSMTPServerPort).Value = Mes.Port
                .Item(cdoSMTPServer).Value = Mes.Server
                .Item(cdoSendUserName).Value = Mes.Username
                .Item(cdoSendPassword).Value = Mes.Password
                .Item(cdoSMTPAuthenticate).Value = cdoBasic
                .Update
            End With
            .From = Mes.FromEmail
            .To = Mes.CellNum & "@" & Mes.Provider
            .Subject = "Texter message"
            .TextBody = Mes.Message
            .Send
        End With
        RaiseEvent MessageSent
    Else
        Err.Raise 380, , "Invalid message format."
    End If
ErrExit:
    Exit Sub
ErrHandler:
    RaiseEvent SendError(Err.Number, Err.Description)
    Resume ErrExit
End Sub

Public Property Get Server() As String
    Server = Mes.Server
End Property

Public Property Let Server(NewVal As String)
    If NewVal = "" Or Len(NewVal) > 50 Then Err.Raise 380, , "Invalid Server."
    Mes.Server = NewVal
    Valid(Props.Serverprop) = True
End Property

Public Property Get Username() As String
    Username = Mes.Username
End Property

Public Property Let Username(NewVal As String)
    If NewVal = "" Or Len(NewVal) > 50 Then Err.Raise 380, , "Invalid Username."
    Mes.Username = NewVal
    Valid(Props.Usernameprop) = True
End Property

