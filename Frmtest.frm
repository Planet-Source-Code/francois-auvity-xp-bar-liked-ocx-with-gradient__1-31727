VERSION 5.00
Object = "*\A..\..\..\AEMPOR~1\CHENILLE\SEND\XPBarVbp.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin XpBarVbp.XPBar XPBar4 
      Height          =   240
      Left            =   1320
      TabIndex        =   4
      Top             =   960
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   423
      LedColor        =   8388608
   End
   Begin XpBarVbp.XPBar XPBar3 
      Height          =   240
      Left            =   1320
      TabIndex        =   3
      Top             =   720
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   423
      LedColor        =   49344
   End
   Begin XpBarVbp.XPBar XPBar2 
      Height          =   240
      Left            =   1320
      TabIndex        =   2
      Top             =   480
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   423
      LedColor        =   128
   End
   Begin XpBarVbp.XPBar XPBar1 
      Height          =   240
      Left            =   1320
      TabIndex        =   1
      Top             =   240
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   423
      LedColor        =   4210752
   End
   Begin VB.CommandButton CmdStop 
      Caption         =   "Stop"
      Height          =   615
      Left            =   7560
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function


'
Sub Gradient(TheObject As Object, Redval&, Greenval&, Blueval&, TopToBottom As Boolean)
    'TheObject can be any object that supports the Line method (like forms and pictures).
    'Redval, Greenval, and Blueval are the Red, Green, and Blue starting values from 0 to 255.
    'TopToBottom determines whether the gradient will draw down or up.
    Dim Step%, Reps%, FillTop%, FillLeft%, FillRight%, FillBottom%, HColor$
    'This will create 63 steps in the gradient. This looks smooth on 16-bit and 24-bit color.
    'You can change this, but be careful. You can do some strange-looking stuff with it...
    Step = (TheObject.Height / 50)
    'This tells it whether to start on the top or the bottom and adjusts variables accordingly.
    If TopToBottom = True Then FillTop = 0 Else FillTop = TheObject.Height - Step
    FillLeft = 0
    FillRight = TheObject.Width
    FillBottom = FillTop + Step
    'If you changed the number of steps, change the number of reps to match it.
    'If you don't, the gradient will look all funny.
    For Reps = 1 To 50
        'This draws the colored bar.
        
        TheObject.Line (FillLeft, FillTop)-(FillRight, FillBottom), RGB(Redval, Greenval, Blueval), BF
        'This decreases the RGB values to darken the color.
        'Lower the value for "squished" gradients. Raise it for incomplete gradients.
        'Also, if you change the number of steps, you will need to change this number.
        If Reps > 25 Then
            Redval = Redval - 3
            Greenval = Greenval - 3
            Blueval = Blueval - 3
        Else
            Redval = Redval + 3
            Greenval = Greenval + 3
            Blueval = Blueval + 3
        End If
'        If Reps = 25 Then
'           TheObject.Line (FillLeft, FillTop)-(FillRight, 409), vbWhite, BF
'
'        End If
        'This prevents the RGB values from becoming negative, which causes a runtime error.
        If Redval <= 0 Then Redval = 0
        If Greenval <= 0 Then Greenval = 0
        If Blueval <= 0 Then Blueval = 0
        'More top or bottom stuff; Moves to next bar.
        If TopToBottom = True Then FillTop = FillBottom Else FillTop = FillTop - Step
        FillBottom = FillTop + Step
    Next

End Sub



Private Sub CmdStop_Click()
XPBar1.Interval = 0

End Sub

Private Sub Command1_Click()

End Sub


