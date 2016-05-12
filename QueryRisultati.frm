VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmQueryRisultati 
   Caption         =   "Risultati della query"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11550
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   11550
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdIndicatori 
      Caption         =   "&Indicatori..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   1
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmdChiudi 
      Caption         =   "&Chiudi"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Top             =   6360
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid msgQueryRis 
      Height          =   5175
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   9128
      _Version        =   393216
      Rows            =   9
      Cols            =   4
      FillStyle       =   1
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.Label Label3 
      Caption         =   "Record selezionati:"
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblRisul 
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
      Left            =   8880
      TabIndex        =   4
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label label1 
      Caption         =   "RISULTATI DELLA QUERY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmQueryRisultati"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChiudi_Click()
  udtFormsLoad.Qrisultati = False
  MDIFd.mnuQueryModifica.Enabled = False
  MDIFd.mnuGraficoIndicatori.Enabled = False
  MDIFd.tlbBarraStrumenti.Buttons(3).Enabled = False
  MDIFd.tlbBarraStrumenti.Buttons(4).Enabled = False
  frmQueryRisultati.Hide
  Unload frmQueryRisultati
End Sub
Private Sub cmdIndicatori_Click()
   frmIndicatori.Show
End Sub
Private Sub Form_Activate()
  udtFormsLoad.Qrisultati = True
  MDIFd.mnuQueryModifica.Enabled = True
  MDIFd.mnuGraficoIndicatori.Enabled = True
  MDIFd.tlbBarraStrumenti.Buttons(3).Enabled = True
  MDIFd.tlbBarraStrumenti.Buttons(4).Enabled = True
End Sub
Private Sub Form_Load()
  udtFormsLoad.Qrisultati = True
  MDIFd.mnuQueryModifica.Enabled = True
  MDIFd.mnuGraficoIndicatori.Enabled = True
  MDIFd.tlbBarraStrumenti.Buttons(3).Enabled = True
  MDIFd.tlbBarraStrumenti.Buttons(4).Enabled = True
  On Error GoTo Errori
  MDIFd.sbrStato.SimpleText = recQuery.RecordCount & " records selezionati: caricamento dei dati nella griglia in corso . . ."
  MDIFd.sbrStato.Refresh
  CaricaGriglia_msgQueryRis
  MDIFd.sbrStato.Visible = False
  MousePointer = 0
  NumeroStocks
Errori:
  Select Case Err.Number
    Case 30006
      o = MsgBox("Memoria insufficiente per visualizzare l'intera query", vbCritical)
      Exit Sub
  End Select
End Sub
Private Function fncIndGriglia(r As Long, c As Integer) As Long
  fncIndGriglia = c + msgQueryRis.Cols * r
End Function
Private Sub CaricaGriglia_msgQueryRis()
  Dim i As Long
  Dim j As Integer
  msgQueryRis.Rows = recQuery.RecordCount + 1
  msgQueryRis.Cols = frmQueryCampi.lstSel.ListCount + 3
  For j = 0 To 1
    msgQueryRis.ColWidth(j) = 1600
  Next
  For j = 2 To frmQueryCampi.lstSel.ListCount + 2
    msgQueryRis.ColWidth(j) = 1300
  Next
  msgQueryRis.Width = 1600 * 7
  lblRisul.Caption = (recQuery.RecordCount)
  msgQueryRis.Row = 0
  msgQueryRis.Col = 1
  msgQueryRis.Text = " Nome stock"
  msgQueryRis.Col = 2
  msgQueryRis.Text = "Data quotazione"
  For i = 1 To UBound(udtDynCampi)
    msgQueryRis.Col = i + 2
    msgQueryRis.Text = udtDynCampi(i).Nome
  Next
  i = 0
  intIndex = 0
  MDIFd.pbrAvanz.Visible = True
  Do Until recQuery.EOF = True
    i = i + 1
    msgQueryRis.TextArray(fncIndGriglia(i, 0)) = "Record " & i
    For j = 0 To frmQueryCampi.lstSel.ListCount + 1
      If j = 1 And frmQueryCampi.lstSel.ListCount > 0 Then j = 2
      If recQuery(j) <> "" Then
        msgQueryRis.TextArray(fncIndGriglia(i, j + 1)) = Format(recQuery(j), "0.0000")
      End If
    Next
    msgQueryRis.TextArray(fncIndGriglia(i, 2)) = Format(recQuery(1), "dd/mm/yyyy")
    MDIFd.pbrAvanz.Value = recQuery.PercentPosition
    recQuery.MoveNext
  Loop
  MDIFd.pbrAvanz.Visible = False
End Sub
Private Sub NumeroStocks()
  Dim strStock As String
  Dim intI As Integer
  Dim intNumber As Integer
  recQuery.MoveFirst
  intI = 0
  intNumber = 0
  Do Until recQuery.EOF = True
    intNumber = intNumber + 1
    strStock = recQuery(0)
    recQuery.MoveNext
    If recQuery.EOF = False Then
      If strStock <> recQuery(0) Then
        intI = intI + 1
        ReDim Preserve udtDynChart(1 To intI)
        udtDynChart(intI).Nome = strStock
        udtDynChart(intI).Numero = intNumber
      End If
    End If
  Loop
  intI = intI + 1
  ReDim Preserve udtDynChart(1 To intI)
  udtDynChart(intI).Nome = strStock
  udtDynChart(intI).Numero = intNumber
End Sub
