VERSION 5.00
Begin VB.Form frmAI 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AI Sample - Lost 0 Times"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrSlowStuff 
      Interval        =   1000
      Left            =   90
      Top             =   510
   End
   Begin VB.Timer tmrAI 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   90
      Top             =   90
   End
   Begin VB.Image imgA 
      Height          =   480
      Left            =   15
      Picture         =   "frmAI.frx":0000
      Top             =   0
      Width           =   480
   End
   Begin VB.Image imgB 
      Height          =   480
      Left            =   4170
      Picture         =   "frmAI.frx":0442
      Top             =   2700
      Width           =   480
   End
End
Attribute VB_Name = "frmAI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim AIMode As VbMsgBoxResult, ChasedLastPos As XYPos, LoseFlag As Boolean, TimesLost As Integer
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case Is = vbKeyLeft
            If imgA.Left <= 0 Then Exit Sub
            imgA.Left = imgA.Left - 10
        Case Is = vbKeyRight
            If imgA.Left >= frmAI.ScaleWidth - 40 Then Exit Sub
            imgA.Left = imgA.Left + 10
        Case Is = vbKeyUp
            If imgA.Top <= 0 Then Exit Sub
            imgA.Top = imgA.Top - 10
        Case Is = vbKeyDown
            If imgA.Top >= frmAI.ScaleHeight - 40 Then Exit Sub
            imgA.Top = imgA.Top + 10
    End Select
End Sub

Private Sub Form_Load()
    AIFollowOn = MsgBox("Press yes to use AIFollow, no to use AIChase.", vbYesNo, "Select One")
    tmrAI.Enabled = True
End Sub


Private Sub tmrAI_Timer()
    If AIMode = vbYes Then
        LoseFlag = AIFollow(ChasedLastPos, imgA, imgB, 10)
    Else
        LoseFlag = AIChase(imgA, imgB, 10)
    End If
End Sub


Private Sub tmrSlowStuff_Timer()
    frmAI.Cls
    If LoseFlag = True Then
        TimesLost = TimesLost + 1
        frmAI.Caption = "AI Sample - Lost " & TimesLost & " Times"
        frmAI.CurrentX = 96
        frmAI.CurrentY = 80
        frmAI.Print "Caught!"
    End If
End Sub


