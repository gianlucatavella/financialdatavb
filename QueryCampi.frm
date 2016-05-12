VERSION 5.00
Begin VB.Form frmQueryCampi 
   Caption         =   "Impostazione query: campi da visualizzare"
   ClientHeight    =   7380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   11625
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdPrecedente 
      Caption         =   "<<  &Precedente"
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
      Left            =   7320
      TabIndex        =   8
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmdSuccessivo 
      Caption         =   "&Successivo  >>"
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
      TabIndex        =   7
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "&Annulla"
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
      TabIndex        =   9
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleziona i campi da visualizzare nella query"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8775
      Begin VB.ListBox lstSel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         Left            =   5280
         TabIndex        =   6
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CommandButton cmdAgg1 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   2
         ToolTipText     =   "Aggiunge un campo"
         Top             =   1440
         Width           =   495
      End
      Begin VB.CommandButton cmdAggT 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         ToolTipText     =   "Aggiunge tutti i campi"
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmdRimT 
         Caption         =   "<<"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   4
         ToolTipText     =   "Rimuove tutti i campi"
         Top             =   2880
         Width           =   495
      End
      Begin VB.CommandButton cmdRim1 
         Caption         =   "<"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   5
         ToolTipText     =   "Rimuove un campo"
         Top             =   3480
         Width           =   495
      End
      Begin VB.ListBox lstCampi 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2985
         Left            =   360
         TabIndex        =   1
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label label1 
         Caption         =   "Campi disponibili"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   11
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Campi selezionati"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4800
         TabIndex        =   10
         Top             =   600
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmQueryCampi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recCampi As Recordset
Dim i As Integer
Private Sub cmdAnnulla_Click()
  ScaricaFormsAttivi
End Sub
Private Sub cmdPrecedente_Click()
  frmQueryCampi.Hide
End Sub
Private Sub cmdSuccessivo_Click()
  Dim strCampo As String
  Dim j As Integer
  strSQL = ""
  strSQL = "SELECT Nomi.nome, Prezzi.data "
  ReDim udtDynCampi(0)
  If lstSel.ListCount <> 0 Then
    ReDim udtDynCampi(1 To lstSel.ListCount)
    For i = 0 To lstSel.ListCount - 1
      strCampo = lstSel.List(i)
      For j = 1 To recCampi.RecordCount
        If strCampo = strDynCampi(j, 2) Then
           strSQL = strSQL & ", " & strDynCampi(j, 1)
           udtDynCampi(i + 1).Nome = strDynCampi(j, 1)
           udtDynCampi(i + 1).Descrizione = strDynCampi(j, 2)
           udtDynCampi(i + 1).Field = i + 2
           Exit For
        End If
      Next
    Next
  End If
  strSQL = strSQL & ", PrimaryK FROM Nomi INNER JOIN Prezzi ON Nomi.codice = Prezzi.codice "
  frmQueryDate.Show
End Sub
Private Sub Form_Activate()
  udtFormsLoad.Qcampi = True
  MDIFd.sbrStato.Visible = True
  MDIFd.sbrStato.SimpleText = "Per selezionare il campo desiderato usare i pulsanti o fare doppio clic sul campo"
  MDIFd.sbrStato.Refresh
End Sub
Private Sub Form_Load()
  udtFormsLoad.Qcampi = True
  Dim strSQLCampi As String
  strSQLCampi = "SELECT Campo, Descrizione FROM Campi ORDER BY Descrizione"
  Set recCampi = dbFinData.OpenRecordset(strSQLCampi, dbOpenSnapshot, dbReadOnly)
  ReDim strDynCampi(1 To recCampi.RecordCount, 1 To 2)
  recCampi.MoveFirst
  For i = 1 To recCampi.RecordCount
    strDynCampi(i, 1) = recCampi(0)
    strDynCampi(i, 2) = recCampi(1)
    recCampi.MoveNext
  Next
  For i = 1 To recCampi.RecordCount
    lstCampi.AddItem strDynCampi(i, 2)
  Next
End Sub
Private Sub cmdAggT_Click()
  lstSel.Clear
  lstCampi.Clear
  cmdAgg1.Enabled = False
  cmdAggT.Enabled = False
  cmdRim1.Enabled = True
  cmdRimT.Enabled = True
  For i = 1 To recCampi.RecordCount
    lstSel.AddItem strDynCampi(i, 2)
  Next
End Sub
Private Sub cmdRimT_Click()
  lstSel.Clear
  lstCampi.Clear
  cmdAgg1.Enabled = True
  cmdAggT.Enabled = True
  cmdRim1.Enabled = False
  cmdRimT.Enabled = False
  For i = 1 To recCampi.RecordCount
    lstCampi.AddItem strDynCampi(i, 2)
  Next
End Sub
Private Sub cmdRim1_Click()
  If lstSel.ListIndex >= 0 Then
    cmdAgg1.Enabled = True
    cmdAggT.Enabled = True
    lstCampi.AddItem lstSel.Text
    lstSel.RemoveItem lstSel.ListIndex
  End If
  If lstSel.ListCount = 0 Then
    cmdRim1.Enabled = False
    cmdRimT.Enabled = False
  Else
    cmdRim1.Enabled = True
    cmdRimT.Enabled = True
  End If
End Sub
Private Sub cmdAgg1_Click()
  If lstCampi.ListIndex >= 0 Then
    cmdRim1.Enabled = True
    cmdRimT.Enabled = True
    lstSel.AddItem lstCampi.Text
    lstCampi.RemoveItem lstCampi.ListIndex
  End If
  If lstCampi.ListCount = 0 Then
    cmdAgg1.Enabled = False
    cmdAggT.Enabled = False
  Else
    cmdAgg1.Enabled = True
    cmdAggT.Enabled = True
  End If
End Sub
Private Sub lstSel_DblClick()
  cmdRim1.Value = True
End Sub
Private Sub lstCampi_DblClick()
  cmdAgg1.Value = True
End Sub
