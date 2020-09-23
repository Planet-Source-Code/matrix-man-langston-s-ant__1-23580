VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   ScaleHeight     =   4800
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   2280
      Max             =   200
      TabIndex        =   5
      Top             =   4440
      Width           =   3855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "End"
      Height          =   195
      Left            =   6600
      TabIndex        =   3
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Stop"
      Height          =   195
      Left            =   6600
      TabIndex        =   2
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   195
      Left            =   6600
      TabIndex        =   1
      Top             =   4080
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3975
      Left            =   0
      ScaleHeight     =   263
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   535
      TabIndex        =   0
      Top             =   0
      Width           =   8055
   End
   Begin VB.Label Label3 
      Caption         =   "Moves:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   3480
      TabIndex        =   6
      Top             =   4080
      Width           =   45
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   4440
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'+-------------------------------------------------------+
'|      Visual Basic implementation of Langston's Ant.   |
'|                                                       |
'|  By: Matrix Man                                       |
'|                                                       |
'+-------------------------------------------------------+
Dim RunOk As Boolean

Private Sub Command1_Click()
    'Start
    Picture1.Cls
    RunOk = True
    RunDemo
End Sub

Private Sub Command2_Click()
    'Stop
    RunOk = False
End Sub

Private Sub Command3_Click()
    'end
    End
End Sub



Private Sub RunDemo()
    Dim CurX As Long, CurY As Long, MoveCounter As Long
    Dim NoiseX As Long, NoiseY As Long
    Dim Directn As Integer, NoiseLoop As Integer
    Dim ThisSq As Long, RandLimit As Long, RandHalf As Long

'Calculate the start position - at the centre
CurX = Picture1.ScaleWidth / 2
CurY = Picture1.ScaleHeight / 2
Directn = Int(4 * Rnd - 1) 'Decide on a direction to face
'now draw in any "noise" level set by the scroll bar
If HScroll1.Value > 0 Then
    RandLimit = 1000 / Screen.TwipsPerPixelY
    RandHalf = RandLimit / 2
    Randomize
    For NoiseLoop = 1 To HScroll1.Value
        NoiseX = Int(((CurX + RandHalf) - (CurX - RandHalf) + 1) * Rnd + (CurX - RandHalf))
        NoiseY = Int(((CurY + RandHalf) - (CurY - RandHalf) + 1) * Rnd + (CurY - RandHalf))
        Picture1.PSet (NoiseX, NoiseY), QBColor(0)
    Next NoiseLoop
End If
'Read the colour of the pixel at the start position
ThisSq = Picture1.Point(CurX, CurY)
Do While RunOk
    Select Case ThisSq
        Case QBColor(0) 'it is black
            Picture1.PSet (CurX, CurY), QBColor(15) 'set this pixel white
            'then turn left
            Directn = Directn - 1
            If Directn < 1 Then
                Directn = 4
            End If
        Case Else 'it should be white
            Picture1.PSet (CurX, CurY), QBColor(0) 'set this pixel black
            'then turn right
            Directn = Directn + 1
            If Directn > 4 Then
                Directn = 1
            End If
        End Select
        Picture1.Refresh
        Select Case Directn
            Case 1 'Up
                CurY = CurY - 1
            Case 2 'Right
                CurX = CurX + 1
            Case 3 'down
                CurY = CurY + 1
            Case 4 'left
                CurX = CurX - 1
        End Select
        ThisSq = Picture1.Point(CurX, CurY) 'read the colour of the next position
        MoveCounter = MoveCounter + 1
        Label1.Caption = Format$(MoveCounter, "#,###,##0")
        Label1.Refresh
        DoEvents 'Gives you a chance to click the stop button
        If CurX = 0 Or CurX = Picture1.ScaleWidth Or CurY = 0 Or CurY = Picture1.ScaleHeight Then
            'rather than scroll the view we come to a stop
            RunOk = False
        End If
    Loop
End Sub

Private Sub HScroll1_Change()
Label2.Caption = HScroll1.Value
End Sub
