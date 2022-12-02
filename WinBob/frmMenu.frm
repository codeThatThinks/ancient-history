VERSION 5.00
Begin VB.Form frmMenu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "WinBob"
   ClientHeight    =   7215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   ControlBox      =   0   'False
   Icon            =   "frmMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   481
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   641
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer frmMenuWalkTime 
      Left            =   9120
      Top             =   120
   End
   Begin VB.PictureBox frmMenuInstr 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1425
      ScaleWidth      =   9585
      TabIndex        =   0
      Top             =   5760
      Width           =   9615
      Begin VB.Image frmMenuInstHide 
         Height          =   210
         Left            =   9240
         Picture         =   "frmMenu.frx":0ABA
         Top             =   120
         Width           =   240
      End
      Begin VB.Label frmMenuKeyEnterLabel 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Open Door"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   4
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label frmMenuKeyRightLabel 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Move Right"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4200
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label frmMenuKeyLeftLabel 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Move Left"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.Image frmMenuKeyEnter 
         Height          =   600
         Left            =   5880
         Picture         =   "frmMenu.frx":0D9C
         Top             =   480
         Width           =   885
      End
      Begin VB.Image frmMenuKeyRight 
         Height          =   645
         Left            =   3600
         Picture         =   "frmMenu.frx":29FE
         Top             =   480
         Width           =   555
      End
      Begin VB.Image frmMenuKeyLeft 
         Height          =   600
         Left            =   1320
         Picture         =   "frmMenu.frx":3D10
         Top             =   480
         Width           =   555
      End
      Begin VB.Label frmMenuInstLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "How To Navigate The Menu:"
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   1
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.Label frmMenuStats 
      Caption         =   "0"
      Height          =   255
      Left            =   0
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "224"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label frmMenuGuyTurn 
      Caption         =   "0"
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label frmMenuGuyWalk 
      Caption         =   "0"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Image frmMenuGuy 
      Height          =   1500
      Left            =   3360
      Picture         =   "frmMenu.frx":4ED2
      Top             =   3960
      Width           =   750
   End
   Begin VB.Image frmMenuBack 
      Height          =   7215
      Left            =   0
      Picture         =   "frmMenu.frx":503F
      Top             =   0
      Width           =   9615
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyLeft Then
        frmMenuGuy.Picture = LoadPicture("images\menuguy.gif")
        frmMenuGuy.Left = frmMenuGuy.Left - 4
        frmMenuGuyTurn.Caption = "0"
        frmMenuWalkTime.Interval = 550
        Label1.Caption = frmMenuGuy.Left
    End If
    If KeyCode = vbKeyRight Then
        frmMenuGuy.Picture = LoadPicture("images\menuguy2.gif")
        frmMenuGuy.Left = frmMenuGuy.Left + 4
        frmMenuGuyTurn.Caption = "1"
        frmMenuWalkTime.Interval = 550
        Label1.Caption = frmMenuGuy.Left
    End If
    If KeyCode = vbKeyReturn Then
        If frmMenuStats.Caption = "0" Then
            If frmMenuGuy.Left >= 19 And frmMenuGuy.Left <= 49 Then
                frmMenuStats.Caption = "1"
                frmMenuBack.Picture = LoadPicture("images\play.bmp")
                frmMenuGuy.Left = 19
                Exit Sub
            End If
            If frmMenuGuy.Left >= 390 And frmMenuGuy.Left <= 422 Then
                frmMenuStats.Caption = "2"
                frmMenuBack.Picture = LoadPicture("images\stats.bmp")
                frmMenuGuy.Left = 19
                Exit Sub
            End If
            If frmMenuGuy.Left >= 486 And frmMenuGuy.Left <= 520 Then
                End
            End If
        End If
        If frmMenuStats.Caption = "1" Then
            If frmMenuGuy.Left >= 19 And frmMenuGuy.Left <= 49 Then
                frmMenuStats.Caption = "0"
                frmMenuBack.Picture = LoadPicture("images\menu.bmp")
                frmMenuGuy.Left = 19
                Exit Sub
            End If
            If frmMenuGuy.Left >= 204 And frmMenuGuy.Left <= 325 Then

            End If
            If frmMenuGuy.Left >= 407 And frmMenuGuy.Left <= 541 Then

            End If
        End If
        If frmMenuStats.Caption = "2" Then
            If frmMenuGuy.Left >= 19 And frmMenuGuy.Left <= 49 Then
                frmMenuStats.Caption = "0"
                frmMenuBack.Picture = LoadPicture("images\menu.bmp")
                frmMenuGuy.Left = 390
            End If
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    frmMenuWalkTime.Interval = "0"
End Sub
Private Sub frmMenuInstHide_Click()
    '456 down
    '384 up
    If frmMenuInstr.Top = 384 Then
        frmMenuInstr.Top = 456
        frmMenuInstHide.Picture = LoadPicture("images\menuhide1.bmp")
        Exit Sub
    End If
    If frmMenuInstr.Top = 456 Then
        frmMenuInstr.Top = 384
        frmMenuInstHide.Picture = LoadPicture("images\menuhide.bmp")
        Exit Sub
    End If
End Sub

Private Sub frmMenuInstHide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If frmMenuInstr.Top = 384 Then
        frmMenuInstHide.Picture = LoadPicture("images\menuhide-push.bmp")
    End If
    If frmMenuInstr.Top = 456 Then
        frmMenuInstHide.Picture = LoadPicture("images\menuhide1-push.bmp")
    End If
End Sub

Private Sub frmMenuInstHide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If frmMenuInstr.Top = 384 Then
        frmMenuInstHide.Picture = LoadPicture("images\menuhide.bmp")
    End If
    If frmMenuInstr.Top = 456 Then
        frmMenuInstHide.Picture = LoadPicture("images\menuhide1.bmp")
    End If
End Sub

Private Sub frmMenuWalkTime_Timer()
    If frmMenuGuyTurn = "0" Then
        If frmMenuGuyWalk = "0" Then
            frmMenuGuyWalk = "1"
            frmMenuGuy.Picture = LoadPicture("images\menuguy3.gif")
            Exit Sub
        End If
        If frmMenuGuyWalk = "1" Then
            frmMenuGuyWalk = "0"
            frmMenuGuy.Picture = LoadPicture("images\menuguy.gif")
            Exit Sub
        End If
        Exit Sub
    End If
    If frmMenuGuyTurn = "1" Then
        If frmMenuGuyWalk = "0" Then
            frmMenuGuyWalk = "1"
            frmMenuGuy.Picture = LoadPicture("images\menuguy4.gif")
            Exit Sub
        End If
        If frmMenuGuyWalk = "1" Then
            frmMenuGuyWalk = "0"
            frmMenuGuy.Picture = LoadPicture("images\menuguy2.gif")
            Exit Sub
        End If
        Exit Sub
    End If
End Sub
