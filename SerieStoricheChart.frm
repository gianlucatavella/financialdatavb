VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmSerieStoricheChart 
   Caption         =   "Grafico della serie storica"
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7290
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin MSChart20Lib.MSChart mscSerie 
      Height          =   7095
      Left            =   0
      OleObjectBlob   =   "SerieStoricheChart.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   11775
   End
End
Attribute VB_Name = "frmSerieStoricheChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
  MousePointer = vbDefault
  MDIFd.mnuGraficoChiudi.Enabled = True
  udtFormsLoad.SstoricheChart = True
End Sub
Private Sub Form_Load()
  MousePointer = vbDefault
  MDIFd.mnuGraficoChiudi.Enabled = True
  udtFormsLoad.SstoricheChart = True
End Sub

