VERSION 5.00
Begin VB.UserControl XPBar 
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   690
   ScaleWidth      =   4800
   ToolboxBitmap   =   "FraXP.ctx":0000
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   960
      Top             =   480
   End
   Begin VB.PictureBox Led 
      BackColor       =   &H00FFFFC0&
      Height          =   240
      Index           =   1
      Left            =   0
      ScaleHeight     =   180
      ScaleWidth      =   135
      TabIndex        =   0
      Top             =   0
      Width           =   200
   End
End
Attribute VB_Name = "XPBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Const LedWidth = 200
Const Interv = 10
Dim i As Integer
Dim blInit As Boolean

Dim TabLed() As PictureBox
Dim NbLed As Integer
Dim noResize As Boolean

Dim RedCo&
Dim GreenCo&
Dim BlueCo&

Dim def_BackColor As Long
'Valeurs de propri�t�s par d�faut:
Const m_def_Appearance = 0
Const m_def_LedColor = &HFFFFC0
Const m_def_BorderStyle = 0
Const m_def_Interval = 50

Const m_def_NbLedOn = 4

'Variables de propri�t�s:
Dim m_Appearance As Integer
Dim m_LedColor As OLE_COLOR
Dim m_BorderStyle As Integer

Dim m_Interval As Long



Private Sub ClearAll()
   For i = 1 To NbLed
    
    TabLed(i).BackColor = def_BackColor
    
   Next

End Sub

Sub Reverse()
Static FirstOn As Integer
Static cptNbLedon As Integer

Static fp As Boolean

Dim x As Integer
Dim LastOn As Integer

On Error Resume Next

If Not fp Then   '
    FirstOn = 1 - m_def_NbLedOn
    fp = True
End If
If FirstOn > NbLed Then
   FirstOn = 1 - m_def_NbLedOn
End If

LastOn = FirstOn + m_def_NbLedOn
FirstOn = FirstOn + 1


       
For x = 1 To NbLed
    Select Case x
    
         
     Case Is < FirstOn
         TabLed(x).BackColor = def_BackColor
        
        
     Case FirstOn To LastOn
        Gradient TabLed(x), RedCo, GreenCo, BlueCo, 1
    
        
      Case Is > LastOn
        TabLed(x).BackColor = def_BackColor

    End Select

Next

If m_Interval = 0 Then Call ClearAll
End Sub

Private Sub TranslateColor(ByVal oClr As OLE_COLOR, Optional hPal As Long = 0)
    ' Convert Automation color to Windows color
    Dim lrgb As Long
    If OleTranslateColor(oClr, hPal, lrgb) Then
        lrgb = CLR_INVALID
    End If
    If lrgb <> CLR_INVALID Then
       GreenCo = (lrgb And &HFF00&) \ &H100
       RedCo = (lrgb And &HFF&)
       BlueCo = (lrgb And &HFF0000) \ &H10000
    End If
    
End Sub


Private Sub Gradient(TheObject As Object, Redval&, Greenval&, Blueval&, TopToBottom As Boolean)
Static Part As Integer
    Dim Step%, Reps%, FillTop%, FillLeft%, FillRight%, FillBottom%, HColor$
    Dim StepW%
    Step = (TheObject.Height / 100)
    If TopToBottom = True Then FillTop = 0 Else FillTop = TheObject.Height - Step
    FillLeft = 0
    FillRight = TheObject.Width
    FillBottom = FillTop + Step
    
    Part = Part + 1
    Select Case Part
    
        Case 1: StepW = 44
        Case 2: StepW = 50
        Case 3: StepW = 56
    End Select

    For Reps = 1 To 100
        
        TheObject.Line (FillLeft, FillTop)-(FillRight, FillBottom), RGB(Redval, Greenval, Blueval), BF
        If Reps > StepW Then
            Redval = Redval - 3
            Greenval = Greenval - 3
            Blueval = Blueval - 3
        Else
            Redval = Redval + 3
            Greenval = Greenval + 3
            Blueval = Blueval + 3
        End If
        If Redval <= 0 Then Redval = 0
        If Greenval <= 0 Then Greenval = 0
        If Blueval <= 0 Then Blueval = 0
        If TopToBottom = True Then FillTop = FillBottom Else FillTop = FillTop - Step
        FillBottom = FillTop + Step
    Next
If Part = 3 Then
   Part = 0
End If
End Sub





Sub GestionTaille()
Dim i As Integer
On Error Resume Next

UserControl.Height = Led(1).Height

NbLed = (UserControl.Width / (LedWidth + Interv))
'noResize = True
'UserControl.Width = (LedWidth + Interv) * NbLed + Interv
'noResize = False
ReDim TabLed(1 To NbLed)
Set TabLed(1) = Led(1)

For i = 2 To NbLed
    Load Led(i)
    Set TabLed(i) = Led(i)
    TabLed(i).Left = TabLed(i - 1).Left + Interv + LedWidth
    TabLed(i).Visible = True
Next

End Sub

Private Sub TestUserMode()
   If Ambient.UserMode Then
      Timer1.Enabled = True
      blInit = True
   Else
      Timer1.Enabled = False
   End If

End Sub

Private Sub Timer1_Timer()
Static FirstOn As Integer
Static cptNbLedon As Integer

Static fp As Boolean

Dim x As Integer
Dim LastOn As Integer

On Error Resume Next

If Not fp Then   '
    FirstOn = 1 - m_def_NbLedOn
    fp = True
End If

If FirstOn > NbLed Then
   FirstOn = 1 - m_def_NbLedOn
End If

LastOn = FirstOn + m_def_NbLedOn
FirstOn = FirstOn + 1


       
For x = 1 To NbLed
    Select Case x
    
         
     Case Is < FirstOn
         TabLed(x).BackColor = def_BackColor
        
        
     Case FirstOn To LastOn
        Gradient TabLed(x), RedCo, GreenCo, BlueCo, 1
    
        
      Case Is > LastOn
        TabLed(x).BackColor = def_BackColor

    End Select

Next

If m_Interval = 0 Then Call ClearAll
End Sub

Private Sub UserControl_Initialize()

Call GestionTaille

End Sub


Private Sub UserControl_Resize()
'If Not noResize Then
    Call GestionTaille
'End If
End Sub


'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENT�ES SUIVANTES!
'MemberInfo=7,0,0,0
Public Property Get Appearance() As Integer
Attribute Appearance.VB_Description = "Renvoie ou d�finit si un objet appara�t ou non en 3D au moment de l'ex�cution."
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENT�ES SUIVANTES!
'MemberInfo=8,0,0,0
Public Property Get LedColor() As OLE_COLOR
Attribute LedColor.VB_Description = "Renvoie ou d�finit la couleur d'arri�re-plan utilis�e pour afficher le texte et les graphiques d'un objet."
    LedColor = m_LedColor
End Property

Public Property Let LedColor(ByVal New_LedColor As OLE_COLOR)
    m_LedColor = New_LedColor
    PropertyChanged "LedColor"
    ' Call TranslateColor(m_LedColor)
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENT�ES SUIVANTES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Renvoie ou d�finit le style de la bordure d'un objet."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'ATTENTION! NE SUPPRIMEZ PAS OU NE MODIFIEZ PAS LES LIGNES COMMENT�ES SUIVANTES!
'MappingInfo=Timer1,Timer1,-1,Interval
Public Property Get Interval() As Long
Attribute Interval.VB_Description = "Renvoie ou d�finit le nombre de millisecondes entre les appels � un �v�nement Timer du contr�le Timer."
    Interval = Timer1.Interval
End Property

Public Property Let Interval(ByVal New_Interval As Long)
    m_Interval = New_Interval
    Timer1.Interval = m_Interval
    If m_Interval = 0 Then
       Call ClearAll
    End If
    PropertyChanged "Interval"
End Property

'Initialiser les propri�t�s pour le contr�le utilisateur
Private Sub UserControl_InitProperties()
    m_Appearance = m_def_Appearance
    m_LedColor = m_def_LedColor
    m_BorderStyle = m_def_BorderStyle
    m_Interval = m_def_Interval
    def_BackColor = UserControl.BackColor
    
    Call TranslateColor(m_def_LedColor)
    Call ClearAll
    
    Call TestUserMode

End Sub
'Charger les valeurs des propri�t�s � partir du stockage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_LedColor = PropBag.ReadProperty("LedColor", m_def_LedColor)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Interval = PropBag.ReadProperty("Interval", m_def_Interval)
    Timer1.Interval = m_Interval
    Call TranslateColor(m_LedColor)
    def_BackColor = UserControl.BackColor
    Call ClearAll
    Call TestUserMode
End Sub


'�crire les valeurs des propri�t�s dans le stockage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("LedColor", m_LedColor, m_def_LedColor)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Interval", m_Interval, m_def_Interval)
End Sub

