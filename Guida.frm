VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmGuida 
   Caption         =   "Guida di Financial Data"
   ClientHeight    =   5745
   ClientLeft      =   1995
   ClientTop       =   2085
   ClientWidth     =   8370
   Icon            =   "Guida.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8370
   Begin VB.TextBox txtGrafici 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Text            =   "Guida.frx":0442
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.TextBox txtComputazioni 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "Guida.frx":05FF
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.TextBox txtInterrogazione 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "Guida.frx":081A
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.TextBox txtApertura 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "Guida.frx":09E2
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.TextBox txtVuoto 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   120
      Width           =   5175
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   5490
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Selezionare l'argomento desiderato cliccandovi sopra"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            TextSave        =   ""
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtPresentazione 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Guida.frx":0ABF
      Top             =   120
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Label lblSerie 
      Caption         =   "Grafici sulle serie storiche"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MouseIcon       =   "Guida.frx":0CEE
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   4440
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblApertura 
      Caption         =   "Apertura della connessione"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MouseIcon       =   "Guida.frx":1130
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1800
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblInterrogazione 
      Caption         =   "Interrogazione della base di dati"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MouseIcon       =   "Guida.frx":1572
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   2640
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblPresentazione 
      Caption         =   "Presentazione"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      MouseIcon       =   "Guida.frx":19B4
      MousePointer    =   99  'Custom
      TabIndex        =   3
      Top             =   960
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblComputazioni 
      Caption         =   "Applicazione di computazioni specifiche sui record ottenuti e relativi grafici"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MouseIcon       =   "Guida.frx":1DF6
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   3480
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   2760
      Y1              =   5520
      Y2              =   0
   End
   Begin VB.Label Label1 
      Caption         =   "      ARGOMENTI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmGuida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  txtVuoto.Visible = True
  txtPresentazione.Visible = False
  txtInterrogazione.Visible = False
  txtComputazioni.Visible = False
  txtGrafici.Visible = False
  txtApertura.Visible = False
End Sub
Private Sub lblApertura_Click()
  txtVuoto.Visible = False
  txtPresentazione.Visible = False
  txtInterrogazione.Visible = False
  txtComputazioni.Visible = False
  txtGrafici.Visible = False
  txtApertura.Visible = True
End Sub
Private Sub lblComputazioni_Click()
  txtVuoto.Visible = False
  txtPresentazione.Visible = False
  txtApertura.Visible = False
  txtInterrogazione.Visible = False
  txtGrafici.Visible = False
  txtComputazioni.Visible = True
End Sub
Private Sub lblInterrogazione_Click()
  txtVuoto.Visible = False
  txtPresentazione.Visible = False
  txtApertura.Visible = False
  txtComputazioni.Visible = False
  txtGrafici.Visible = False
  txtInterrogazione.Visible = True
End Sub
Private Sub lblPresentazione_Click()
  txtVuoto.Visible = False
  txtApertura.Visible = False
  txtInterrogazione.Visible = False
  txtComputazioni.Visible = False
  txtGrafici.Visible = False
  txtPresentazione.Visible = True
End Sub
Private Sub lblSerie_Click()
  txtVuoto.Visible = False
  txtApertura.Visible = False
  txtInterrogazione.Visible = False
  txtComputazioni.Visible = False
  txtPresentazione.Visible = False
  txtGrafici.Visible = True
End Sub
