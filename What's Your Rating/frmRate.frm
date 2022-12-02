VERSION 5.00
Begin VB.Form frmRate 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "What's Your Rating?"
   ClientHeight    =   1590
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "frmRate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton frmRateFigure 
      Caption         =   "Click Here To Find Out!"
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   840
      Width           =   2295
   End
   Begin VB.TextBox frmRateEntry 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   2280
      TabIndex        =   7
      Text            =   "Enter your rate here"
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label frmRateExcellent 
      BackColor       =   &H00008000&
      Caption         =   "Above 34 = Excellent"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label frmRateAboveAverage 
      BackColor       =   &H0000FF00&
      Caption         =   "28 to 34 = Above average "
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label frmRateAverage 
      BackColor       =   &H0000FFFF&
      Caption         =   "23 to 27 = Average "
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label frmRateBelowAverage 
      BackColor       =   &H000080FF&
      Caption         =   "16 to 22 = Below average "
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Label frmRatePoor 
      BackColor       =   &H000000FF&
      Caption         =   "Below 16 = Poor"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label frmRateLvl 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label frmRateTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "What's Your Rating?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   3975
   End
   Begin VB.Image frmRateLogo 
      Height          =   345
      Left            =   120
      Picture         =   "frmRate.frx":0442
      Top             =   0
      Width           =   360
   End
   Begin VB.Menu btn_exit 
      Caption         =   "&Exit"
   End
   Begin VB.Menu btn_about 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmRate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btn_about_Click()
    MsgBox "This program was made with VB6.0 and is avalible at www.geocities.com/glen_knw. You can do what ever you want with this program."
End Sub

Private Sub btn_exit_Click()
    End
End Sub

Private Sub frmRateFigure_Click()
    Select Case Val(frmRateEntry.Text)
        Case Is < 16
            frmRateLvl.Caption = "Poor"
            
        Case Is <= 22
            frmRateLvl.Caption = "Below Average"
            
        Case Is <= 27
            frmRateLvl.Caption = "Average"
            
        Case Is <= 34
            frmRateLvl.Caption = "Above Average"
            
        Case Is > 34
            frmRateLvl.Caption = "Excellent"
    End Select
End Sub
