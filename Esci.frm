VERSION 5.00
Begin VB.Form frmEsci 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "           Esci da Financial Data"
   ClientHeight    =   1770
   ClientLeft      =   3915
   ClientTop       =   3495
   ClientWidth     =   3720
   Icon            =   "Esci.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "Annulla"
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
      Left            =   2160
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
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
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Fine della sessione di lavoro."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Esci.frx":0442
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmEsci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnnulla_Click()
  frmEsci.Hide
  Unload frmEsci
End Sub
Private Sub cmdOk_Click()
  frmEsci.Hide
  Unload frmEsci
  ScaricaFormsAttivi
  End
End Sub
