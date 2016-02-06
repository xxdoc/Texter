VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10080
   Icon            =   "frmSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Textbox 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   1680
      TabIndex        =   32
      Top             =   5160
      Width           =   3000
   End
   Begin VB.TextBox Provider 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   6840
      TabIndex        =   20
      Top             =   2280
      Width           =   3000
   End
   Begin VB.TextBox Provider 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   6840
      TabIndex        =   16
      Top             =   1800
      Width           =   3000
   End
   Begin VB.TextBox Provider 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   6840
      TabIndex        =   12
      Top             =   1320
      Width           =   3000
   End
   Begin VB.CommandButton butCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7080
      TabIndex        =   33
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton butClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   8520
      TabIndex        =   34
      Top             =   5040
      Width           =   1335
   End
   Begin VB.TextBox Textbox 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   1680
      TabIndex        =   23
      Top             =   3240
      Width           =   1320
   End
   Begin VB.TextBox Textbox 
      Alignment       =   2  'Center
      Height          =   285
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1680
      TabIndex        =   31
      Top             =   4680
      Width           =   3000
   End
   Begin VB.TextBox Textbox 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   26
      Top             =   3720
      Width           =   3000
   End
   Begin VB.TextBox Textbox 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   1680
      TabIndex        =   29
      Top             =   4200
      Width           =   3000
   End
   Begin VB.TextBox Cell 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   3600
      TabIndex        =   7
      Top             =   840
      Width           =   3000
   End
   Begin VB.TextBox Cell 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   3600
      TabIndex        =   11
      Top             =   1320
      Width           =   3000
   End
   Begin VB.TextBox Cell 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   3600
      TabIndex        =   15
      Top             =   1800
      Width           =   3000
   End
   Begin VB.TextBox Cell 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   3600
      TabIndex        =   19
      Top             =   2280
      Width           =   3000
   End
   Begin VB.TextBox tbName 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Top             =   840
      Width           =   2520
   End
   Begin VB.TextBox tbName 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   1
      Left            =   720
      TabIndex        =   10
      Top             =   1320
      Width           =   2520
   End
   Begin VB.TextBox tbName 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   2
      Left            =   720
      TabIndex        =   14
      Top             =   1800
      Width           =   2520
   End
   Begin VB.TextBox tbName 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   3
      Left            =   720
      TabIndex        =   18
      Top             =   2280
      Width           =   2520
   End
   Begin VB.CheckBox ck 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   255
   End
   Begin VB.CheckBox ck 
      Height          =   285
      Index           =   1
      Left            =   240
      TabIndex        =   9
      Top             =   1320
      Width           =   255
   End
   Begin VB.CheckBox ck 
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   1800
      Width           =   255
   End
   Begin VB.CheckBox ck 
      Height          =   285
      Index           =   3
      Left            =   240
      TabIndex        =   17
      Top             =   2280
      Width           =   255
   End
   Begin VB.TextBox Provider 
      Alignment       =   2  'Center
      Height          =   285
      Index           =   0
      Left            =   6840
      TabIndex        =   8
      Top             =   840
      Width           =   3000
   End
   Begin VB.Label Label11 
      Caption         =   "From Email"
      Height          =   285
      Left            =   120
      TabIndex        =   35
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label 
      Caption         =   "SMTP Server"
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
      TabIndex        =   21
      Top             =   2880
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Port"
      Height          =   285
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "(ex: SMPTP2GO can use 25)"
      Height          =   285
      Left            =   5040
      TabIndex        =   24
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label10 
      Caption         =   "(ex: smtpcorp.com)"
      Height          =   285
      Left            =   5040
      TabIndex        =   27
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label4 
      Caption         =   "Password"
      Height          =   285
      Left            =   120
      TabIndex        =   30
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Server"
      Height          =   285
      Left            =   120
      TabIndex        =   25
      Top             =   3720
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Username"
      Height          =   285
      Left            =   120
      TabIndex        =   28
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "Cell #'s"
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
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "Name"
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Top             =   480
      Width           =   2520
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "10 digit # including area code"
      Height          =   285
      Left            =   3600
      TabIndex        =   3
      Top             =   480
      Width           =   3000
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      Caption         =   "Select"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   480
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Provider (ex: sms.phoneco.com)"
      Height          =   285
      Left            =   6840
      TabIndex        =   4
      Top             =   480
      Width           =   3000
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Loading As Boolean

Private Sub butCancel_Click()
    Load
    butCancel.Enabled = False
    butClose.Caption = "Close"
End Sub

Private Sub butClose_Click()
    On Error GoTo ErrHandler
    If butClose.Caption = "Save" Then
        Save
        Edited False
    Else
        Unload Me
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmSetup", "butClose_Click", Err.Description
     Resume ErrExit
End Sub

Private Sub Cell_Change(Index As Integer)
    Edited
End Sub

Private Sub ck_Click(Index As Integer)
    Edited
End Sub

Private Sub Edited(Optional NewVal As Boolean = True)
    On Error GoTo ErrHandler
    If Not Loading Then
        If NewVal Then
            butCancel.Enabled = True
            butClose.Caption = "Save"
        Else
            butCancel.Enabled = False
            butClose.Caption = "Close"
        End If
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmSetup", "Edited", Err.Description
     Resume ErrExit
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHandler
    Select Case KeyCode
        Case 38
            'up arrow
            SendKeys ("+{tab}")
        Case 40
            'down arrow
            SendKeys ("{tab}")
    End Select
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSetup", "Form_KeyDown", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrHandler
    If KeyAscii = 13 Then
        'enter
        SendKeys ("{tab}")
        KeyAscii = 0
    End If
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSetup", "Form_KeyPress", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_Load()
    AD.LoadFormData Me
    Load
End Sub

Private Sub Form_Unload(Cancel As Integer)
    AD.SaveFormData Me
End Sub

Private Sub Load()
    Dim C As Long
    On Error GoTo ErrHandler
    Loading = True
    For C = 0 To 3
        ck(C).Value = Val(AD.AppData("Select" & C))
        tbName(C).Text = AD.AppData("Name" & C)
        Cell(C).Text = AD.AppData("Cell" & C)
        Provider(C).Text = AD.AppData("Provider" & C)
    Next C
    TextBox(0).Text = AD.AppData("Port")
    TextBox(1).Text = AD.AppData("Server")
    TextBox(2).Text = AD.AppData("Username")
    TextBox(3).Text = AD.AppData("Password")
    TextBox(4).Text = AD.AppData("From")
    Loading = False
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmSetup", "Load", Err.Description
     Resume ErrExit
End Sub

Private Sub Provider_Change(Index As Integer)
    Edited
End Sub

Private Sub Save()
    Dim C As Long
    On Error GoTo ErrHandler
    For C = 0 To 3
        AD.AppData("Select" & C) = ck(C)
        AD.AppData("Name" & C) = tbName(C)
        AD.AppData("Cell" & C) = Cell(C)
        AD.AppData("Provider" & C) = Provider(C)
    Next C
    AD.AppData("Port") = TextBox(0).Text
    AD.AppData("Server") = TextBox(1).Text
    AD.AppData("Username") = TextBox(2).Text
    AD.AppData("Password") = TextBox(3).Text
    AD.AppData("From") = TextBox(4).Text
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
     AD.DisplayError Err.Number, "frmSetup", "Save", Err.Description
     Resume ErrExit
End Sub

Private Sub tbName_Change(Index As Integer)
    Edited
End Sub

Private Sub Textbox_Change(Index As Integer)
    Edited
End Sub

Private Sub textbox_GotFocus(Index As Integer)
    On Error GoTo ErrHandler
    TextBox(Index).SelStart = 0
    TextBox(Index).SelLength = Len(TextBox(Index).Text)
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "frmSetup", "textbox_GotFocus", Err.Description
    Resume ErrExit
End Sub

