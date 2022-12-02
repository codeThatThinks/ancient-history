VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "WinBob"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5250
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   250
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   4680
      Top             =   120
   End
   Begin VB.Image frmSplashTruck2 
      Height          =   450
      Left            =   -1500
      Picture         =   "frmSplash.frx":0ABA
      Top             =   3000
      Width           =   450
   End
   Begin VB.Image frmSplashTruck 
      Height          =   450
      Left            =   -975
      Picture         =   "frmSplash.frx":0B62
      Top             =   3000
      Width           =   450
   End
   Begin VB.Image frmSplashTank 
      Height          =   225
      Left            =   -450
      Picture         =   "frmSplash.frx":0C0A
      Top             =   3240
      Width           =   465
   End
   Begin VB.Label frmSplashVersion 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Version 1.0"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label frmSplashLabel 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "WinBob"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Image frmSplashBackground 
      Height          =   3750
      Left            =   0
      Picture         =   "frmSplash.frx":0CB1
      Top             =   0
      Width           =   6000
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

End Sub

Private Sub Timer1_Timer()
    frmSplashTank.Left = frmSplashTank.Left + 1
    frmSplashTruck.Left = frmSplashTruck.Left + 2
    frmSplashTruck2.Left = frmSplashTruck2.Left + 2
    If frmSplashTank.Left = 182 Then
        Timer1.Interval = 0
        frmSplashTank.Picture = LoadPicture("images\splashtank1.gif")
        frmSplashTank.Top = frmSplashTank.Top - 15
        Dim sec
        sec = Time$
        Do
        Loop Until Time$ <> sec
        frmSplashTank.Picture = LoadPicture("images\splashtank2.gif")
        sec = Time$
        Do
        Loop Until Time$ <> sec
        frmMenu.Show
        frmSplash.Hide
    End If
End Sub
