VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   0  'None
   ClientHeight    =   2085
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3030
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShiftFocus 
      Caption         =   "Dummy"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Form 4 within App"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1695
      Begin VB.OptionButton optBtn 
         Caption         =   "Option 4"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1400
      End
      Begin VB.OptionButton optBtn 
         Caption         =   "Option 3"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1400
      End
      Begin VB.OptionButton optBtn 
         Caption         =   "Option 2"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1400
      End
      Begin VB.OptionButton optBtn 
         Caption         =   "Option 1"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1400
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'!! NOTE: cmdShiftFocus must not be viewable on the screen as
'         its only purpose is to identify when to pass focus
'         back to the parent form

Dim OptPtr As Long
Event MoveFocus()

Private Sub cmdShiftFocus_GotFocus()
    '## Reset tabbing order
    optBtn(OptPtr).SetFocus
    '## Advise parent form that we are releasing focus back
    RaiseEvent MoveFocus
End Sub

Private Sub optBtn_Click(Index As Integer)
    OptPtr = Index
End Sub
