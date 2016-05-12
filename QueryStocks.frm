VERSION 5.00
Begin VB.Form frmQueryStocks 
   AutoRedraw      =   -1  'True
   Caption         =   "Impostazione query: selezione stocks"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   690
   ClientWidth     =   11880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdAnnulla 
      Caption         =   "&Annulla"
      Height          =   615
      Left            =   3840
      TabIndex        =   9
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmdSuccessivo 
      Caption         =   "&Successivo  >>"
      Enabled         =   0   'False
      Height          =   615
      Left            =   9600
      TabIndex        =   8
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleziona fra gli stocks disponibili"
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
      Width           =   7215
      Begin VB.Frame Frame2 
         Caption         =   "Selezione per settore"
         Height          =   1095
         Left            =   1200
         TabIndex        =   12
         Top             =   4440
         Width           =   4815
         Begin VB.ComboBox cmbTipologia 
            Height          =   315
            Left            =   360
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   480
            Width           =   3015
         End
      End
      Begin VB.ListBox lstNomi 
         Height          =   2985
         Left            =   360
         TabIndex        =   1
         Top             =   1200
         Width           =   2415
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
         Left            =   3360
         TabIndex        =   5
         ToolTipText     =   "Rimuove uno stock"
         Top             =   3480
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
         Left            =   3360
         TabIndex        =   4
         ToolTipText     =   "Rimuove tutti gli stock"
         Top             =   2880
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
         Left            =   3360
         TabIndex        =   3
         ToolTipText     =   "Aggiunge tutti gli stock"
         Top             =   2040
         Width           =   495
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
         Left            =   3360
         TabIndex        =   2
         ToolTipText     =   "Aggiunge uno stock"
         Top             =   1440
         Width           =   495
      End
      Begin VB.ListBox lstSel 
         Height          =   2985
         Left            =   4440
         TabIndex        =   6
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label2 
         Caption         =   "Stock selezionati"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   11
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label label1 
         Caption         =   "Stock disponibili"
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
         TabIndex        =   10
         Top             =   600
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmQueryStocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim recStocksDisp As Recordset
Dim recStocksSect As Recordset
Dim i As Integer
Private Sub cmdAggT_Click()
  lstSel.Clear
  lstNomi.Clear
  cmdAgg1.Enabled = False
  cmdAggT.Enabled = False
  cmdRim1.Enabled = True
  cmdRimT.Enabled = True
  recStocksDisp.MoveFirst
  For i = 1 To recStocksDisp.RecordCount
    lstSel.AddItem recStocksDisp(0)
    recStocksDisp.MoveNext
  Next
  cmdSuccessivo.Enabled = True
End Sub
Private Sub cmdAnnulla_Click()
  ScaricaFormsAttivi
End Sub
Private Sub cmdSuccessivo_Click()
  If lstNomi.ListCount <> 0 Then
    strSQLstocks = " WHERE ("
    For i = 0 To lstSel.ListCount - 1
      strSQLstocks = strSQLstocks & " nome = '" & lstSel.List(i) & "' OR "
    Next
    strSQLstocks = Mid(strSQLstocks, 1, Len(strSQLstocks) - 3)
    strSQLstocks = strSQLstocks & ")"
  Else
    strSQLstocks = ""
  End If
  frmQueryCampi.Show
End Sub
Private Sub cmdRimT_Click()
  lstSel.Clear
  lstNomi.Clear
  cmdAgg1.Enabled = True
  cmdAggT.Enabled = True
  cmdRim1.Enabled = False
  cmdRimT.Enabled = False
  recStocksDisp.MoveFirst
  For i = 1 To recStocksDisp.RecordCount
    lstNomi.AddItem recStocksDisp(0)
    recStocksDisp.MoveNext
  Next
  cmdSuccessivo.Enabled = False
End Sub
Private Sub cmdRim1_Click()
    If lstSel.ListIndex >= 0 Then
      cmdAgg1.Enabled = True
      cmdAggT.Enabled = True
      lstNomi.AddItem lstSel.Text
      lstSel.RemoveItem lstSel.ListIndex
    End If
    If lstSel.ListCount = 0 Then
      cmdRim1.Enabled = False
      cmdRimT.Enabled = False
      cmdSuccessivo.Enabled = False
    Else
      cmdRim1.Enabled = True
      cmdRimT.Enabled = True
    End If
End Sub
Private Sub cmdAgg1_Click()
  If lstNomi.ListIndex >= 0 Then
    cmdRim1.Enabled = True
    cmdRimT.Enabled = True
    lstSel.AddItem lstNomi.Text
    lstNomi.RemoveItem lstNomi.ListIndex
  End If
  If lstNomi.ListCount = 0 Then
    cmdAgg1.Enabled = False
    cmdAggT.Enabled = False
  Else
    cmdAgg1.Enabled = True
    cmdAggT.Enabled = True
    If lstSel.ListCount <> 0 Then
      cmdSuccessivo.Enabled = True
    End If
  End If
End Sub
Private Sub Form_Activate()
  udtFormsLoad.Qstocks = True
  MDIFd.sbrStato.Visible = True
  MDIFd.sbrStato.SimpleText = "Per selezionare lo stock desiderato usare i pulsanti o fare doppio clic sullo stock"
  MDIFd.sbrStato.Refresh
End Sub
Private Sub lstSel_DblClick()
  cmdRim1.Value = True
End Sub
Private Sub lstnomi_DblClick()
  cmdAgg1.Value = True
End Sub
Private Sub cmbTipologia_Click()
  If cmbTipologia.Text <> "" Then
    lstSel.Clear
    lstNomi.Clear
    recStocksDisp.MoveFirst
    For i = 1 To recStocksDisp.RecordCount
      If recStocksDisp(1) = cmbTipologia.Text Then
        lstSel.AddItem recStocksDisp(0)
      Else
        lstNomi.AddItem recStocksDisp(0)
      End If
      recStocksDisp.MoveNext
    Next
    cmdSuccessivo.Enabled = True
    cmdRim1.Enabled = True
    cmdRimT.Enabled = True
  End If
End Sub
Private Sub Form_Load()
  udtFormsLoad.Qstocks = True
  Dim strSQLStocksDisp As String
  Dim strSQLStocksSect As String
  strSQLStocksDisp = "SELECT Nome, Settore FROM Nomi ORDER BY Nome"
  Set recStocksDisp = dbFinData.OpenRecordset(strSQLStocksDisp, dbOpenSnapshot, dbReadOnly)
  strSQLStocksSect = "SELECT DISTINCT settore FROM Nomi "
  Set recStocksSect = dbFinData.OpenRecordset(strSQLStocksSect, dbOpenSnapshot, dbReadOnly)
  recStocksDisp.MoveFirst
  For i = 1 To recStocksDisp.RecordCount
    lstNomi.AddItem recStocksDisp(0)
    recStocksDisp.MoveNext
  Next
  recStocksSect.MoveFirst
  For i = 1 To recStocksSect.RecordCount
    cmbTipologia.AddItem recStocksSect(0)
    recStocksSect.MoveNext
  Next
End Sub


