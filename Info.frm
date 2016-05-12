VERSION 5.00
Begin VB.Form frmInfo 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informazioni su Financial Data"
   ClientHeight    =   4935
   ClientLeft      =   1470
   ClientTop       =   1665
   ClientWidth     =   8820
   ControlBox      =   0   'False
   Icon            =   "Info.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8820
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Image Image6 
      Height          =   525
      Left            =   5040
      Picture         =   "Info.frx":000C
      Top             =   0
      Width           =   405
   End
   Begin VB.Image Image5 
      Height          =   525
      Left            =   4680
      Picture         =   "Info.frx":0BCA
      Top             =   1200
      Width           =   405
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   2280
      Picture         =   "Info.frx":1788
      Top             =   720
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   1560
      Picture         =   "Info.frx":1BCA
      Top             =   360
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1440
      Picture         =   "Info.frx":200C
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label Label1 
      Caption         =   $"Info.frx":244E
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   3720
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   2850
      Left            =   240
      Picture         =   "Info.frx":28EB
      Top             =   120
      Width           =   2865
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOk_Click()
  Unload frmInfo
End Sub

