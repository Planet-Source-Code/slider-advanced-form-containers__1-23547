VERSION 5.00
Begin VB.Form frmParent 
   Caption         =   "Form1"
   ClientHeight    =   2535
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2535
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Show Options (Page 2)"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   120
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   2400
      ScaleHeight     =   1215
      ScaleWidth      =   2055
      TabIndex        =   3
      Top             =   600
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1215
      ScaleWidth      =   2055
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Main Host Form"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmParent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const Demo1 = 0        '## Set to 1 to view 'in-App-Form' Demo
                        '   or '0' for 'in-process-DLL' Demo
'## Child form hWnd
Private mlF2Ptr As Long
Private mlF3Ptr As Long
Private mlF4Ptr As Long

'## Child's parent hWnd
Private mlP2Ptr As Long
Private mlP3Ptr As Long
Private mlP4Ptr As Long

'## Picture Container Next item hWnd in tab order
Private mlFocus1Ptr As Long
Private mlFocus2Ptr As Long

'## Child form declarations
Private WithEvents oF2 As Form2
Attribute oF2.VB_VarHelpID = -1
Private WithEvents oF3 As Form3
Attribute oF3.VB_VarHelpID = -1

#If Demo1 Then
    Private WithEvents oF4 As Form4
Attribute oF4.VB_VarHelpID = -1
#Else
    '## a class is used as an passthrough process as
    '   forms are not exposed from DLLs
    Private WithEvents oF4 As pDllForm.cForm1
Attribute oF4.VB_VarHelpID = -1
#End If

Private Sub Check1_Click()
    '## Switching disblayed forms
    Select Case Check1.Value
        Case 0
            '## Release current child form and hide
            pReleaseContainer mlF4Ptr, mlP4Ptr
            oF4.Visible = False
            '## Assign new child form and show
            mlP3Ptr = pAssignContainer(mlF3Ptr, Picture2.hWnd)
            oF3.Visible = True
        Case 1
            pReleaseContainer mlF3Ptr, mlP3Ptr
            oF3.Visible = False
            mlP4Ptr = pAssignContainer(mlF4Ptr, Picture2.hWnd)
            oF4.Visible = True
    End Select
End Sub

Private Sub Form_Load()

    Set oF2 = New Form2
    Set oF3 = New Form3

#If Demo1 Then
    Set oF4 = New Form4
#Else
    Set oF4 = New pDllForm.cForm1
#End If

    mlF2Ptr = oF2.hWnd
    mlF3Ptr = oF3.hWnd
    mlF4Ptr = oF4.hWnd

    pSetWinStyle mlF2Ptr
    pSetWinStyle mlF3Ptr
    pSetWinStyle mlF4Ptr

    mlP2Ptr = pAssignContainer(mlF2Ptr, Picture1.hWnd)
    mlP3Ptr = pAssignContainer(mlF3Ptr, Picture2.hWnd)

    oF2.Visible = True
    oF3.Visible = True
    
    mlFocus1Ptr = Picture2.hWnd
    mlFocus2Ptr = Check1.hWnd

End Sub

Private Sub Form_Unload(Cancel As Integer)

    pReleaseContainer mlF2Ptr, mlP2Ptr
    Select Case Check1.Value
        Case 0: pReleaseContainer mlF3Ptr, mlP3Ptr
        Case 1: pReleaseContainer mlF4Ptr, mlP4Ptr
    End Select
    Unload oF2
    Unload oF3
    Set oF2 = Nothing
    Set oF3 = Nothing

#If Demo1 Then
    Unload oF4
#End If
    Set oF4 = Nothing

End Sub

Private Sub pSetWinStyle(hWnd As Long)
    Dim dwStyle As Long
    '## set the form's window style
    dwStyle = GetWindowLong(hWnd, GWL_STYLE)
    dwStyle = dwStyle Or WS_CHILD
    SetWindowLong hWnd, GWL_STYLE, dwStyle
End Sub

'## Assign a child form to the parent container
Private Function pAssignContainer(lFrmPtr As Long, lPicPtr As Long) As Long
    pAssignContainer = SetParent(lFrmPtr, lPicPtr)
    '## move the form flush inside the picturebox frame
    Call SetWindowPos(lFrmPtr, HWND_TOP, 0, 0, 0, 0, SWP_NOSIZE)
End Function

'## Must be called to release the child form from the parent container
Private Sub pReleaseContainer(lFrmPtr As Long, lParentPtr As Long)
    SetParent lFrmPtr, lParentPtr
End Sub

Private Sub oF2_MoveFocus()
    '## Do not use .SetFocus method as we are shifting focus between
    '   two forms
    Call winSetFocus(mlFocus1Ptr) '## Manually select next object in tab
                                  '   order on parent form
End Sub

Private Sub oF3_MoveFocus()
    Call winSetFocus(mlFocus2Ptr)
End Sub

Private Sub oF4_MoveFocus()
    Call winSetFocus(mlFocus2Ptr)
End Sub

Private Sub Picture1_GotFocus()
    Call winSetFocus(mlF2Ptr)   '## Shift focus to child form
End Sub

Private Sub Picture2_GotFocus()
    Call winSetFocus(IIf(Check1.Value = 0, mlF3Ptr, mlF4Ptr))
End Sub
