VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "A 2D Rotation Demo v1.0"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar ScrollCosine 
      Height          =   345
      LargeChange     =   45
      Left            =   690
      Max             =   180
      SmallChange     =   5
      TabIndex        =   11
      Top             =   4860
      Width           =   2805
   End
   Begin VB.HScrollBar ScrollSine 
      Height          =   345
      LargeChange     =   45
      Left            =   690
      Max             =   180
      SmallChange     =   5
      TabIndex        =   9
      Top             =   2520
      Width           =   2805
   End
   Begin VB.CommandButton btnRotate 
      Caption         =   "Rotate"
      Enabled         =   0   'False
      Height          =   345
      Left            =   7590
      TabIndex        =   8
      Top             =   4860
      Width           =   1245
   End
   Begin VB.Timer TimerRotate 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7050
      Top             =   4950
   End
   Begin VB.CommandButton btnCosine 
      Caption         =   "Cosine ( n )"
      Enabled         =   0   'False
      Height          =   345
      Left            =   3570
      TabIndex        =   7
      Top             =   4860
      Width           =   1245
   End
   Begin VB.CommandButton btnSIN 
      Caption         =   "Sine ( n )"
      Height          =   345
      Left            =   3570
      TabIndex        =   6
      Top             =   2520
      Width           =   1245
   End
   Begin VB.PictureBox pictCircle 
      AutoRedraw      =   -1  'True
      Height          =   3795
      Left            =   5040
      ScaleHeight     =   3735
      ScaleWidth      =   3735
      TabIndex        =   5
      Top             =   1020
      Width           =   3795
   End
   Begin VB.PictureBox pictCosine 
      AutoRedraw      =   -1  'True
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4635
      TabIndex        =   4
      Top             =   3360
      Width           =   4695
   End
   Begin VB.PictureBox pictSine 
      AutoRedraw      =   -1  'True
      Height          =   1455
      Left            =   120
      ScaleHeight     =   1395
      ScaleWidth      =   4635
      TabIndex        =   3
      Top             =   1020
      Width           =   4695
   End
   Begin VB.PictureBox pictTop 
      Align           =   1  'Align Top
      BackColor       =   &H80000005&
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   8865
      TabIndex        =   0
      Top             =   0
      Width           =   8925
      Begin VB.Image Image1 
         Height          =   480
         Left            =   60
         Picture         =   "Form1.frx":0442
         Top             =   135
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A 2D Rotation Demo v1.0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   600
         TabIndex        =   2
         Top             =   120
         Width           =   4125
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Demonstrates how to use SIN() and COS() to draw circles and/or rotate objects about an origin."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   600
         TabIndex        =   1
         Top             =   495
         Width           =   8280
      End
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Offset"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   4935
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Offset"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   2595
      Width           =   420
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_RotationAngle As Single
Private Const g_sngPIDivideBy180 As Single = 0.01745329

Public Function ConvertDeg2Rad(Degress As Single) As Single
    
    ' Converts Degrees to Radians
    ConvertDeg2Rad = Degress * (g_sngPIDivideBy180)
    
End Function

Private Sub btnCosine_Click()

    Dim sngAngleInDegrees As Single
    Dim sngAngleInRadians As Single
    Dim sngX As Single
    Dim sngY As Single
    
    
    ' =================================================================================
    ' Make the drawing area of the picturebox (pictCosine) equal to the following values.
    ' X coordinates from 0 to 360
    ' Y coordinates from +1 to -1
    ' =================================================================================
    Me.pictCosine.ScaleLeft = 0
    Me.pictCosine.ScaleTop = -1
    Me.pictCosine.ScaleHeight = 2 ' ie. from -1 to +1 = 2.
    Me.pictCosine.ScaleWidth = 360
        
    ' ===================================================
    ' Erase the picture box and set some drawing options.
    ' ===================================================
    Me.pictCosine.BackColor = vbBlack
    Me.pictCosine.ForeColor = RGB(255, 0, 0)
    Me.pictCosine.AutoRedraw = True
    Me.pictCosine.Cls
    
    
    ' Draw the centre line.
    Me.pictCosine.DrawWidth = 1
    Me.pictCosine.Line (0, 0)-(360, 0), RGB(0, 64, 64)
    
    ' =============================
    ' Perform a loop from 0 to 360.
    ' =============================
    Me.pictCosine.DrawWidth = 2
    For sngAngleInDegrees = 0 To 360 Step 1
    
        ' Convert the degrees to radians.
        sngAngleInRadians = ConvertDeg2Rad(sngAngleInDegrees + Me.ScrollCosine.Value)
        
        ' Get the new pixel locations.
        sngX = sngAngleInDegrees
        sngY = Cos(sngAngleInRadians)
        
        Me.pictCosine.PSet (sngX, sngY)
        
    Next sngAngleInDegrees
    
    Me.btnRotate.Enabled = True
    
End Sub

Private Sub btnRotate_Click()

    ' =================================================================================
    ' Make the drawing area of the picturebox (pictSine) equal to the following values.
    ' X coordinates from 0 to 360
    ' Y coordinates from +1 to -1
    ' =================================================================================
    Me.pictCircle.ScaleLeft = -1
    Me.pictCircle.ScaleTop = -1
    Me.pictCircle.ScaleHeight = 2 ' ie. from -1 to +1 = 2.
    Me.pictCircle.ScaleWidth = 2
    
    ' Clear picturebox.
    Me.pictCircle.BackColor = vbBlack
    Me.pictCircle.Cls
    
    ' Reset the rotation animation.
    m_RotationAngle = 0
    Me.TimerRotate.Enabled = True
    
End Sub

Private Sub btnSIN_Click()

    Dim sngAngleInDegrees As Single
    Dim sngAngleInRadians As Single
    Dim sngX As Single
    Dim sngY As Single
    
    
    ' =================================================================================
    ' Make the drawing area of the picturebox (pictSine) equal to the following values.
    ' X coordinates from 0 to 360
    ' Y coordinates from +1 to -1
    ' =================================================================================
    Me.pictSine.ScaleLeft = 0
    Me.pictSine.ScaleTop = -1
    Me.pictSine.ScaleHeight = 2 ' ie. from -1 to +1 = 2.
    Me.pictSine.ScaleWidth = 360
    
    
    ' ===================================================
    ' Erase the picture box and set some drawing options.
    ' ===================================================
    Me.pictSine.BackColor = vbBlack
    Me.pictSine.ForeColor = RGB(0, 255, 0)
    Me.pictSine.AutoRedraw = True
    Me.pictSine.Cls


    ' =====================
    ' Draw the centre line.
    ' =====================
    Me.pictSine.DrawWidth = 1
    Me.pictSine.Line (0, 0)-(360, 0), RGB(0, 64, 64)
    
    
    ' =============================
    ' Perform a loop from 0 to 360.
    ' =============================
    Me.pictSine.DrawWidth = 2
    For sngAngleInDegrees = 0 To 360 Step 1
    
        ' Convert the degrees to radians.
        sngAngleInRadians = ConvertDeg2Rad(sngAngleInDegrees + Me.ScrollSine.Value)
        
        ' Get the new pixel locations.
        sngX = sngAngleInDegrees
        sngY = Sin(sngAngleInRadians)
        
        ' Plot the pixel.
        Me.pictSine.PSet (sngX, sngY)
        
    Next sngAngleInDegrees
    
    Me.btnCosine.Enabled = True
    
End Sub

Private Sub ScrollCosine_Change()
    Call btnCosine_Click
End Sub

Private Sub ScrollCosine_Scroll()
    Call btnCosine_Click
End Sub

Private Sub scrollSine_Change()
    Call btnSIN_Click
End Sub

Private Sub scrollSine_Scroll()
    Call btnSIN_Click
End Sub

Private Sub TimerRotate_Timer()

    Dim sngAngleInDegrees As Single
    Dim sngAngleInRadians As Single
    Dim sngX As Single
    Dim sngY As Single
    
    ' Increment module level variable called 'm_RotationAngle'.
    m_RotationAngle = m_RotationAngle + 1
    If m_RotationAngle > 360 Then
        Me.TimerRotate.Enabled = False
        Exit Sub
    End If
    
    ' Clear picture box
    Me.pictCircle.Cls
    Me.pictCircle.ForeColor = RGB(0, 255, 255)
    Me.pictCircle.DrawWidth = 2
    
    ' =============================
    ' Draw the Circle (in progress)
    ' =============================
    For sngAngleInDegrees = 0 To m_RotationAngle Step 0.25
    
        ' Get the radians (taking into consideration the scrollbar offset value)
        sngAngleInRadians = ConvertDeg2Rad(sngAngleInDegrees + Me.ScrollCosine.Value)
        sngX = Cos(sngAngleInRadians)
        
        ' Get the radians (taking into consideration the scrollbar offset value)
        sngAngleInRadians = ConvertDeg2Rad(sngAngleInDegrees + Me.ScrollSine.Value)
        sngY = Sin(sngAngleInRadians)
        
        ' Plot the point.
        Me.pictCircle.PSet (sngX, sngY)
        
    Next sngAngleInDegrees
    
        
    ' Draw Sine & Cosine parts (on the circle part).
    ' ==============================================
    Me.pictCircle.Line (0, 0)-(sngX, 0), RGB(255, 0, 0)
    Me.pictCircle.Line (0, 0)-(0, sngY), RGB(0, 255, 0)
    
    
    ' Update the position on the original Sine graph.
    ' ===============================================
    Me.pictSine.DrawWidth = 1
    Me.pictSine.AutoRedraw = False
    Me.pictSine.Cls
    Me.pictSine.Line (m_RotationAngle, 0)-(m_RotationAngle, sngY)
    
    
    ' Update the position on the original Sine graph.
    ' ===============================================
    Me.pictCosine.DrawWidth = 1
    Me.pictCosine.AutoRedraw = False
    Me.pictCosine.Cls
    Me.pictCosine.Line (m_RotationAngle, 0)-(m_RotationAngle, sngX)
    
    
    ' Draw a big yellow knob at the end.
    ' ==================================
    Me.pictCircle.DrawWidth = 10
    Me.pictCircle.PSet (sngX, sngY), RGB(255, 255, 0)
                
End Sub
