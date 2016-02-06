Attribute VB_Name = "modMain"
Option Explicit
Public AD As ClsAppData

Sub Main()
    If App.PrevInstance Then End
    Set AD = New ClsAppData
    InitCommonControlsVB
    frmMain.Show
End Sub

