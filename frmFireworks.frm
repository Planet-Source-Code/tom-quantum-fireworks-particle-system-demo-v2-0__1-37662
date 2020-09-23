VERSION 5.00
Begin VB.Form frmFireworks 
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox pctInfo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   120
      ScaleHeight     =   2295
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.Label lblTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Fireworks Demo v2.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label lblDesc 
         BackStyle       =   0  'Transparent
         Caption         =   "Demonstrates the use of a simple particle engine."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   5
         Top             =   360
         Width           =   4335
      End
      Begin VB.Label lblLClick 
         BackStyle       =   0  'Transparent
         Caption         =   "Left-Click: Create a new fireworks explosion"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   4
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label lblRClick 
         BackStyle       =   0  'Transparent
         Caption         =   "Right-Click: Quit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label lblArrows 
         BackStyle       =   0  'Transparent
         Caption         =   "Up/Down arrow keys: Increase/decrease gravity"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   1080
         Width           =   4335
      End
      Begin VB.Label lblGravity 
         BackStyle       =   0  'Transparent
         Caption         =   "GRAVITY ="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   1320
         Width           =   4335
      End
   End
   Begin VB.Timer tmrRepeat 
      Interval        =   3000
      Left            =   4200
      Top             =   2640
   End
   Begin VB.Timer tmrRefresh 
      Interval        =   50
      Left            =   3720
      Top             =   2640
   End
End
Attribute VB_Name = "frmFireworks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
'      FIREWORKS v.2.0
'
'      Demonstrates how to use a 2D particle engine
'-----------------------------------------------------------

Option Explicit

Private Const NUM_PARTICLES = 500       ' Max # of particles
Private GRAVITY As Single               ' Gravity, not listed as a Const so that it can change
Private MIN_X As Long                   ' Screen, minimum X
Private MIN_Y As Long                   ' Screen, minimum Y
Private MAX_X As Long                   ' Screen, maximum X
Private MAX_Y As Long                   ' Screen, maximum Y

' Form_Activate - Do initialization. Can't use Form_Load because Me.Scale won't work then.
Private Sub Form_Activate()
    InitScreen
    InitParticles 0, MAX_Y / 2
    GRAVITY = 1
    lblGravity.Caption = "GRAVITY = " & Format(GRAVITY * 10, "00.0")
End Sub

' Form_KeyUp - Up/Down arrows handle gravity
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyUp
        GRAVITY = GRAVITY + 0.1
    Case vbKeyDown
        GRAVITY = GRAVITY - 0.1
    End Select
    lblGravity.Caption = "GRAVITY = " & Format(GRAVITY * 10, "00.0")
End Sub

' Form_MouseUp - Create a fireworks blast
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        Unload Me
        Exit Sub
    End If
    
    tmrRepeat.Enabled = False
    tmrRepeat.Enabled = True
    InitParticles X, Y
    'Debug.Print X, Y
End Sub

' tmrRefresh_Timer -  Update all particles
Private Sub tmrRefresh_Timer()
    RefreshParticles
End Sub

' InitParticles - Create NUM_PARTICLES particles
Private Sub InitParticles(ByVal X As Long, ByVal Y As Long)
    Dim i As Integer
    Dim p As Particle
    Dim Angle As Single
    Dim Velocity As Single
    Dim ParticleColor As Long
    
    ' This has a side effect of removing all particles on-screen, but this was the only way
    ' to remove particles, as you can't implement a For Each block for arrays so you can
    ' delete unused particles. Any help with this is appreciated!
    ReDim ParticleSys(0)
    Randomize
    
    ' Choose a random color for this batch of particles
    ParticleColor = Int(Rnd * 3)
    For i = 1 To NUM_PARTICLES
        ' Initialize position
        p.X = X
        p.Y = Y
        
        ' Set a random angle and velocity for it to fire in
        Angle = Rnd * 6.28
        Velocity = Rnd * 20 + 1
        
        ' Simple Trig: Cos(Angle) * Constant = Cartesian X-value. Same goes for Sin and Y
        p.VX = Cos(Angle) * Velocity
        p.VY = Sin(Angle) * Velocity
        
        ' Random size
        p.Size = Int(Rnd * 2) + 1
        
        ' Fireworks come in 3 random colors!
        Select Case ParticleColor
        Case 0
            p.Color = vbRed
        Case 1
            p.Color = vbGreen
        Case 2
            p.Color = vbBlue
        End Select
        
        ' Now finally create the particle
        CreateParticle p
    Next i
End Sub

' InitScreen - Scale the screen to 1024x768 for easier processing (and that's my screen res :)
Private Sub InitScreen()
    Dim w As Integer
    Dim h As Integer
    
    Me.BackColor = vbBlack
    w = 1024
    h = 768
    MIN_X = -w / 2
    MIN_Y = 0
    MAX_X = w / 2
    MAX_Y = h
    Me.Scale (MIN_X, MAX_Y)-(MAX_X, MIN_Y)
End Sub

' RefreshParticles - Update all particles' positions and velocities, and then draw them
Private Sub RefreshParticles()
    UpdParticles
    UpdVelocities
    DrawParticles
End Sub

' UpdVelocities - This is where gravity influences the Y velocity of all particles. It
' can arguably be implemented in the ParticleSystem module, but different programs have
' different effects (i.e. wind) so I decided to put this in the form module instead so as to
' keep the module generalized.
Private Sub UpdVelocities()
    On Error Resume Next
    Dim i As Long
    Dim l As Long
    
    l = UBound(ParticleSys)
    
    Randomize
    For i = 1 To l
        With ParticleSys(i)
            ' To simulate different masses of particles, gravity is a bit randomized.
            .VY = .VY - GRAVITY + ((Rnd * GRAVITY) - GRAVITY / 2)
            
            ' Faster particles age faster because of the wind cooling them down
            .Age = .Age + (Sqr(.VX ^ 2 + .VY ^ 2) / 10) + 1
        End With
    Next i
End Sub

' DrawParticles - Simply draw the particles.
Private Sub DrawParticles()
    On Error Resume Next
    Dim i As Long
    Dim l As Long
    Dim Color As Long
    Dim Red As Long
    Dim Green As Long
    Dim Blue As Long
    Dim OX As Single
    Dim OY As Single
    
    l = UBound(ParticleSys)
    
    Cls
    For i = 1 To l
        With ParticleSys(i)
            Red = 255 - (.Age * 3)
            If Red < 0 Then Red = 0
            Green = 255 / (.Age / 5)
            Blue = 255 / (.Age / 3)
            Select Case .Color
            Case vbRed
                Color = RGB(Red, Green, Blue)
            Case vbGreen
                Color = RGB(Green, Red, Blue)
            Case vbBlue
                Color = RGB(Blue, Green, Red)
            End Select
            If (.Age > 1) Then      ' Older particles are drawn with a trail
                OX = .X - .VX
                OY = .Y - .VY
                
                Me.DrawWidth = .Size
                Me.Line (OX, OY)-(.X, .Y), Color
            Else                    ' New particle, simply draw a pixel
                Me.DrawWidth = .Size
                Me.PSet (.X, .Y), Color
            End If
        End With
    Next i
End Sub

' tmrRepeat_Timer - Every tmrRepeat.Interval, make a new fireworks blast
Private Sub tmrRepeat_Timer()
    Dim X As Long
    Dim Y As Long
    
    Randomize
    X = Int(Rnd * (MAX_X - MIN_X)) + MIN_X
    Y = Int(Rnd * MAX_Y)
    'Debug.Print X, Y
    InitParticles X, Y
End Sub
