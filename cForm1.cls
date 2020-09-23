VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents oF1 As Form1
Attribute oF1.VB_VarHelpID = -1

Event MoveFocus()

Private Sub Class_Initialize()
    Set oF1 = New Form1
End Sub

Private Sub Class_Terminate()
    Set oF1 = Nothing
End Sub

Private Sub oF1_MoveFocus()
    RaiseEvent MoveFocus
End Sub

Public Property Get hWnd() As Long
    hWnd = oF1.hWnd
End Property

Public Property Let Visible(State As Boolean)
    oF1.Visible = State
End Property

