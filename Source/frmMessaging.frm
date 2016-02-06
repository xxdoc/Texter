VERSION 5.00
Begin VB.Form frmMessaging 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setup"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Textbox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   1320
   End
   Begin VB.CheckBox CheckToEmail 
      Caption         =   "send message to email"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   21
      Top             =   4680
      Width           =   2895
   End
   Begin VB.CheckBox CheckToCell 
      Caption         =   "send message to cell phone"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   4680
      Width           =   3975
   End
   Begin VB.TextBox Textbox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   7
      Left            =   1800
      TabIndex        =   18
      Top             =   4080
      Width           =   3000
   End
   Begin VB.TextBox Textbox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   6
      Left            =   1800
      TabIndex        =   15
      Top             =   3528
      Width           =   3000
   End
   Begin VB.TextBox Textbox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   5
      Left            =   1800
      TabIndex        =   12
      Top             =   2980
      Width           =   3000
   End
   Begin VB.TextBox Textbox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   4
      Left            =   1800
      TabIndex        =   10
      Top             =   2432
      Width           =   3000
   End
   Begin VB.TextBox Textbox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      IMEMode         =   3  'DISABLE
      Index           =   3
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   1884
      Width           =   3000
   End
   Begin VB.TextBox Textbox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   2
      Left            =   1800
      TabIndex        =   6
      Top             =   1336
      Width           =   3000
   End
   Begin VB.TextBox Textbox 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   1
      Left            =   1800
      TabIndex        =   3
      Top             =   788
      Width           =   3000
   End
   Begin VB.Label Label1 
      Caption         =   "Port"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "(ex: Gmail uses 465)"
      Height          =   315
      Left            =   5160
      TabIndex        =   2
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label12 
      Caption         =   "(ex: sms.phoneco.com)"
      Height          =   315
      Left            =   5160
      TabIndex        =   17
      Top             =   3648
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "(10 digit # including area code)"
      Height          =   315
      Left            =   5160
      TabIndex        =   14
      Top             =   3100
      Width           =   2295
   End
   Begin VB.Label Label10 
      Caption         =   "(ex: smtp.gmail.com)"
      Height          =   315
      Left            =   5160
      TabIndex        =   5
      Top             =   908
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "To email"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   19
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label7 
      Caption         =   "Provider"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   16
      Top             =   3648
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "Cell number"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   3100
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "From email"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   11
      Top             =   2552
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   2004
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   7
      Top             =   1456
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   908
      Width           =   975
   End
End
Attribute VB_Name = "frmMessaging"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const FldCount As Long = 7
Dim CurrentTextBox As Integer
Private modNotifier As clsMessaging

Private Sub CheckToCell_Click()
    On Error GoTo ErrHandler
    modNotifier.CellOn = CheckBoxToBool(CheckToCell)
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    DisplayError Err.Number, "frmMessaging", "CheckToCell_Click", Err.Description
    Resume ErrExit
End Sub

Private Sub CheckToEmail_Click()
    On Error GoTo ErrHandler
    modNotifier.EmailOn = CheckBoxToBool(CheckToEmail)
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    DisplayError Err.Number, "frmMessaging", "CheckToEmail_Click", Err.Description
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
    DisplayError Err.Number, "frmMessaging", "Form_KeyDown", Err.Description
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
    DisplayError Err.Number, "frmMessaging", "Form_KeyPress", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    AD.LoadFormData Me
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    DisplayError Err.Number, "frmMessaging", "Form_Load", Err.Description
    Resume ErrExit
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo ErrHandler
    Textbox_Validate CurrentTextBox, False
    AD.SaveFormData Me
    Set modNotifier = Nothing
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    DisplayError Err.Number, "frmMessaging", "Form_Unload", Err.Description
    Resume ErrExit
End Sub

Private Sub LoadData()
    Dim F As Long
    On Error GoTo ErrHandler
    With modNotifier
        For F = 0 To FldCount - 1
            Select Case F
                Case 0
                    TextBox(F) = .Port
                Case 1
                    TextBox(F) = .Server
                Case 2
                    TextBox(F) = .UserName
                Case 3
                    TextBox(F) = .Password
                Case 4
                    TextBox(F) = .FromEmail
                Case 5
                    TextBox(F) = .CellNumber
                Case 6
                    TextBox(F) = .Provider
                Case 7
                    TextBox(F) = .ToEmail
            End Select
        Next F
        SetCheckBox CheckToCell, .CellOn
        SetCheckBox CheckToEmail, .EmailOn
    End With
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    DisplayError Err.Number, "frmMessaging", "LoadData", Err.Description
    Resume ErrExit
End Sub

Public Sub SetNotifier(Notifier As clsMessaging)
    Set modNotifier = Notifier
    LoadData
End Sub

Private Sub textbox_GotFocus(Index As Integer)
    On Error GoTo ErrHandler
    TextBox(Index).SelStart = 0
    TextBox(Index).SelLength = Len(TextBox(Index).Text)
    CurrentTextBox = Index
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    DisplayError Err.Number, "frmMessaging", "textbox_GotFocus", Err.Description
    Resume ErrExit
End Sub

Private Sub Textbox_Validate(Index As Integer, Cancel As Boolean)
    On Error GoTo ErrHandler
    With modNotifier
        Select Case Index
            Case 0
                .Port = Val(TextBox(Index).Text)
            Case 1
                .Server = TextBox(Index).Text
            Case 2
                .UserName = TextBox(Index).Text
            Case 3
                .Password = TextBox(Index).Text
            Case 4
                .FromEmail = TextBox(Index).Text
            Case 5
                .CellNumber = TextBox(Index).Text
            Case 6
                .Provider = TextBox(Index).Text
            Case 7
                .ToEmail = TextBox(Index).Text
        End Select
    End With
    On Error GoTo 0
ErrExit:
    Exit Sub
ErrHandler:
    DisplayError Err.Number, "frmMessaging", "Textbox_Validate", Err.Description
    Resume ErrExit
End Sub

