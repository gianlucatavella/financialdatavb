VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIndicatoriRis 
   Caption         =   "Indicatori descrittivi tipici del mercato azionario: risultati delle computazioni"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   11640
   Begin VB.CommandButton cmdGrafico 
      Caption         =   "G&rafico"
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
      Height          =   615
      Left            =   9600
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
   Begin MSFlexGridLib.MSFlexGrid msgIndicatoreRis 
      Height          =   5175
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   9128
      _Version        =   393216
      Rows            =   9
      Cols            =   4
      FillStyle       =   1
      MergeCells      =   1
      AllowUserResizing=   3
   End
   Begin VB.Label lblIndicatori 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   4575
   End
End
Attribute VB_Name = "frmIndicatoriRis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChiudi_Click()
  frmIndicatoriRis.Hide
  Unload frmIndicatoriRis
End Sub
Private Sub cmdGrafico_Click()
    intIndexS = 1
    frmIndicatoriChart.Show
End Sub
Private Sub Form_Activate()
  udtFormsLoad.Irisultati = True
  lblIndicatori.Caption = frmIndicatori.cmbIndicatori.Text
End Sub
Private Sub Form_Load()
  udtFormsLoad.Irisultati = True
  MDIFd.sbrStato.Visible = True
  MDIFd.sbrStato.SimpleText = "Caricamento della griglia in corso . . ."
  MDIFd.sbrStato.Refresh
  If frmIndicatori.cmbIndicatori.Text <> "Indice Stocastico" Then
    CaricaGriglia_msgIndicatoreRis
  Else
    CaricaGriglia_msgIndicatoreRisStocastico
  End If
  MDIFd.sbrStato.Visible = False
End Sub
Private Sub CaricaGriglia_msgIndicatoreRis()
  Dim i As Long
  Dim j As Integer
  recQuery.MoveFirst
  msgIndicatoreRis.Rows = recQuery.RecordCount + 1
  msgIndicatoreRis.Cols = 4
  For j = 0 To 1
    msgIndicatoreRis.ColWidth(j) = 1600
  Next
  msgIndicatoreRis.ColWidth(2) = 1300
  msgIndicatoreRis.ColWidth(3) = 1900
  msgIndicatoreRis.Width = 1600 * 5
  With msgIndicatoreRis
    .Row = 0
    .Col = 1
    .Text = " Nome stock"
    .Col = 2
    .Text = "Data quotazione"
    .Col = 3
    .Text = frmIndicatori.cmbIndicatori.Text
  End With
  i = 0
  MDIFd.pbrAvanz.Value = 0
  MDIFd.pbrAvanz.Visible = True
  Do Until recQuery.EOF = True
    msgIndicatoreRis.TextArray(fncIndGriglia(i + 1, 0)) = "Record " & i + 1
    If recQuery(0) <> "" Then
      msgIndicatoreRis.TextArray(fncIndGriglia(i + 1, 1)) = recQuery(0)
    End If
    msgIndicatoreRis.TextArray(fncIndGriglia(i + 1, 2)) = Format(recQuery(1), "dd/mm/yyyy")
    If dblDynIndicatore(i) <> -1 Then
      msgIndicatoreRis.TextArray(fncIndGriglia(i + 1, 3)) = Format(dblDynIndicatore(i), "0.000")
    Else
      msgIndicatoreRis.TextArray(fncIndGriglia(i + 1, 3)) = "Non disponibile"
    End If
    MDIFd.pbrAvanz.Value = recQuery.PercentPosition
    recQuery.MoveNext
    i = i + 1
  Loop
  MDIFd.pbrAvanz.Visible = False
End Sub
Private Function fncIndGriglia(r As Long, c As Integer) As Long
  fncIndGriglia = c + msgIndicatoreRis.Cols * r
End Function
Private Sub CaricaGriglia_msgIndicatoreRisStocastico()
  Dim i As Long
  Dim j As Integer
  recQuery.MoveFirst
  msgIndicatoreRis.Rows = recQuery.RecordCount + 1
  msgIndicatoreRis.Cols = 5
  For j = 0 To 1
    msgIndicatoreRis.ColWidth(j) = 1600
  Next
  For j = 2 To 4
    msgIndicatoreRis.ColWidth(j) = 1300
  Next
  msgIndicatoreRis.Width = 1600 * 5
  With msgIndicatoreRis
    .Row = 0
    .Col = 1
    .Text = " Nome stock"
    .Col = 2
    .Text = "Data quotazione"
    .Col = 3
    .Text = "%K"
    .Col = 4
    .Text = "%D"
  End With
  i = 0
  MDIFd.pbrAvanz.Value = 0
  MDIFd.pbrAvanz.Visible = True
  Do Until recQuery.EOF = True
    msgIndicatoreRis.TextArray(fncIndGriglia(i + 1, 0)) = "Record " & i + 1
    If recQuery(0) <> "" Then
      msgIndicatoreRis.TextArray(fncIndGriglia(i + 1, 1)) = recQuery(0)
    End If
    msgIndicatoreRis.TextArray(fncIndGriglia(i + 1, 2)) = Format(recQuery(1), "dd/mm/yyyy")
    If dblDynStocastico(i).K <> -1 Then
      msgIndicatoreRis.TextArray(fncIndGriglia(i + 1, 3)) = Format(dblDynStocastico(i).K, "0.000")
      If i > 2 Then
        msgIndicatoreRis.TextArray(fncIndGriglia(i + 1, 4)) = Format(dblDynStocastico(i).D, "0.000")
      Else
        msgIndicatoreRis.TextArray(fncIndGriglia(i + 1, 4)) = "Non disponibile"
      End If
    Else
      msgIndicatoreRis.TextArray(fncIndGriglia(i + 1, 3)) = "Non disponibile"
      msgIndicatoreRis.TextArray(fncIndGriglia(i + 1, 4)) = "Non disponibile"
    End If
    MDIFd.pbrAvanz.Value = recQuery.PercentPosition
    recQuery.MoveNext
    i = i + 1
  Loop
  MDIFd.pbrAvanz.Visible = False
End Sub
