VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   ClientHeight    =   2070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3030
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   720
      Width           =   1575
   End
   Begin VB.CommandButton cmdShiftFocus 
      Caption         =   "Dummy"
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "We're in Form1 within Inprocess DLL"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'!! NOTE: cmdShiftFocus must not be viewable on the screen as
'         its only purpose is to identify when to pass focus
'         back to the parent form

Event MoveFocus()

Private Sub cmdShiftFocus_GotFocus()
    '## Reset tabbing order
    Text2.SetFocus
    '## Advise parent form that we are releasing focus back
    RaiseEvent MoveFocus
End Sub

