VERSION 5.00
Begin VB.Form frmSerieStoriche 
   Caption         =   "Serie storiche"
   ClientHeight    =   7425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   11625
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdGrafico 
      Caption         =   "G&rafico"
      Default         =   -1  'True
      Enabled         =   0   'False
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
      TabIndex        =   10
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
      TabIndex        =   11
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Visualizzazione delle serie storiche tramite grafici"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   10695
      Begin VB.ComboBox cmbCampo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3360
         Width           =   3375
      End
      Begin VB.ComboBox cmbTipoGraf 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3360
         Width           =   3375
      End
      Begin VB.ComboBox cmbY2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6840
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox cmbM2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5400
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox cmbD2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2040
         Width           =   735
      End
      Begin VB.ComboBox cmbY1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox cmbM1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2040
         Width           =   1335
      End
      Begin VB.ComboBox cmbD1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2040
         Width           =   735
      End
      Begin VB.ComboBox cmbStocks 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label lblCampo 
         Caption         =   "Campo da descrivere"
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
         Left            =   6240
         TabIndex        =   17
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo di grafico richiesto "
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
         Left            =   1440
         TabIndex        =   16
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Selezione dell'intervallo temporale richiesto"
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
         Left            =   2400
         TabIndex        =   15
         Top             =   1560
         Width           =   4455
      End
      Begin VB.Label Label3 
         Caption         =   "Selezione tra gli stocks disponibili"
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
         Left            =   960
         TabIndex        =   14
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label2 
         Caption         =   "a"
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
         Left            =   4200
         TabIndex        =   13
         Top             =   2160
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Da"
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
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   375
      End
   End
End
Attribute VB_Name = "frmSerieStoriche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strTipi(1 To 5) As String
Dim i As Integer
Dim strMesi(1 To 12) As String
Private Sub cmbCampo_Click()
  If (cmbCampo <> "" Or cmbTipoGraf = strTipi(5)) And cmbD1 <> "" And cmbD2 <> "" And cmbM1 <> "" And cmbM2 <> "" And cmbY1 <> "" And cmbY2 <> "" And cmbStocks <> "" And cmbTipoGraf <> "" Then
    cmdGrafico.Enabled = True
  Else
    cmdGrafico.Enabled = False
  End If
End Sub
Private Sub cmbD1_Click()
  If (cmbCampo <> "" Or cmbTipoGraf = strTipi(5)) And cmbD1 <> "" And cmbD2 <> "" And cmbM1 <> "" And cmbM2 <> "" And cmbY1 <> "" And cmbY2 <> "" And cmbStocks <> "" And cmbTipoGraf <> "" Then
    cmdGrafico.Enabled = True
  Else
    cmdGrafico.Enabled = False
  End If
End Sub
Private Sub cmbD2_Click()
  If (cmbCampo <> "" Or cmbTipoGraf = strTipi(5)) And cmbD1 <> "" And cmbD2 <> "" And cmbM1 <> "" And cmbM2 <> "" And cmbY1 <> "" And cmbY2 <> "" And cmbStocks <> "" And cmbTipoGraf <> "" Then
    cmdGrafico.Enabled = True
  Else
    cmdGrafico.Enabled = False
  End If
End Sub
Private Sub cmbM1_Click()
  If (cmbCampo <> "" Or cmbTipoGraf = strTipi(5)) And cmbD1 <> "" And cmbD2 <> "" And cmbM1 <> "" And cmbM2 <> "" And cmbY1 <> "" And cmbY2 <> "" And cmbStocks <> "" And cmbTipoGraf <> "" Then
    cmdGrafico.Enabled = True
  Else
    cmdGrafico.Enabled = False
  End If
End Sub
Private Sub cmbM2_Click()
  If (cmbCampo <> "" Or cmbTipoGraf = strTipi(5)) And cmbD1 <> "" And cmbD2 <> "" And cmbM1 <> "" And cmbM2 <> "" And cmbY1 <> "" And cmbY2 <> "" And cmbStocks <> "" And cmbTipoGraf <> "" Then
    cmdGrafico.Enabled = True
  Else
    cmdGrafico.Enabled = False
  End If
End Sub
Private Sub cmbStocks_Click()
  If (cmbCampo <> "" Or cmbTipoGraf = strTipi(5)) And cmbD1 <> "" And cmbD2 <> "" And cmbM1 <> "" And cmbM2 <> "" And cmbY1 <> "" And cmbY2 <> "" And cmbStocks <> "" And cmbTipoGraf <> "" Then
    cmdGrafico.Enabled = True
  Else
    cmdGrafico.Enabled = False
  End If
End Sub
Private Sub cmbTipoGraf_Click()
  If (cmbCampo <> "" Or cmbTipoGraf = strTipi(5)) And cmbD1 <> "" And cmbD2 <> "" And cmbM1 <> "" And cmbM2 <> "" And cmbY1 <> "" And cmbY2 <> "" And cmbStocks <> "" And cmbTipoGraf <> "" Then
    cmdGrafico.Enabled = True
  Else
    cmdGrafico.Enabled = False
  End If
  If cmbTipoGraf = strTipi(5) Then
    cmbCampo.Enabled = False
    lblCampo.Enabled = False
  Else
    cmbCampo.Enabled = True
    lblCampo.Enabled = True
  End If
End Sub
Private Sub cmbY1_Click()
  If (cmbCampo <> "" Or cmbTipoGraf = strTipi(5)) And cmbD1 <> "" And cmbD2 <> "" And cmbM1 <> "" And cmbM2 <> "" And cmbY1 <> "" And cmbY2 <> "" And cmbStocks <> "" And cmbTipoGraf <> "" Then
    cmdGrafico.Enabled = True
  Else
    cmdGrafico.Enabled = False
  End If
End Sub
Private Sub cmbY2_Click()
  If (cmbCampo <> "" Or cmbTipoGraf = strTipi(5)) And cmbD1 <> "" And cmbD2 <> "" And cmbM1 <> "" And cmbM2 <> "" And cmbY1 <> "" And cmbY2 <> "" And cmbStocks <> "" And cmbTipoGraf <> "" Then
    cmdGrafico.Enabled = True
  Else
    cmdGrafico.Enabled = False
  End If
End Sub
Private Sub cmdChiudi_Click()
  MDIFd.sbrStato.Visible = False
  ScaricaFormsAttivi
End Sub
Private Sub cmdGrafico_Click()
  Call Grafico
End Sub
Private Sub Form_Activate()
  MousePointer = vbDefault
  udtFormsLoad.Sstoriche = True
  MDIFd.sbrStato.Visible = True
  MDIFd.sbrStato.SimpleText = "Impostare i parametri desiderati"
  MDIFd.sbrStato.Refresh
End Sub
Private Sub Form_Load()
  udtFormsLoad.Sstoriche = True
  MousePointer = vbDefault
  MDIFd.sbrStato.Visible = True
  MDIFd.sbrStato.SimpleText = "Impostare i parametri desiderati"
  MDIFd.sbrStato.Refresh
  Dim strSQLStocksDisp As String
  Dim recStocks As Recordset
  Dim strSQLCampi As String
  Dim recCampi As Recordset
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
    cmbCampo.AddItem strDynCampi(i, 2)
  Next
  strSQLStocksDisp = "SELECT Nome FROM Nomi ORDER BY Nome"
  Set recStocks = dbFinData.OpenRecordset(strSQLStocksDisp, dbOpenSnapshot, dbReadOnly)
  recStocks.MoveFirst
  For i = 1 To recStocks.RecordCount
    cmbStocks.AddItem recStocks(0)
    recStocks.MoveNext
  Next
  strMesi(1) = "Gennaio"
  strMesi(2) = "Febbraio"
  strMesi(3) = "Marzo"
  strMesi(4) = "Aprile"
  strMesi(5) = "Maggio"
  strMesi(6) = "Giugno"
  strMesi(7) = "Luglio"
  strMesi(8) = "Agosto"
  strMesi(9) = "Settembre"
  strMesi(10) = "Ottobre"
  strMesi(11) = "Novembre"
  strMesi(12) = "Dicembre"
  strTipi(1) = "Barre"
  strTipi(2) = "Barre 3D"
  strTipi(3) = "Linee"
  strTipi(4) = "Linee 3D"
  strTipi(5) = "Candle Stick"
  For i = 1 To 12
    cmbM1.AddItem strMesi(i)
    cmbM2.AddItem strMesi(i)
  Next
  For i = 1 To 31
    cmbD1.AddItem Str(i)
    cmbD2.AddItem Str(i)
  Next
  For i = intMin To intMax
    cmbY1.AddItem Str(i)
    cmbY2.AddItem Str(i)
  Next
  For i = 1 To 5
    cmbTipoGraf.AddItem strTipi(i)
  Next
End Sub
Private Sub Grafico()
  On Error GoTo Subscript
  MousePointer = vbHourglass
  MDIFd.sbrStato.SimpleText = "Caricamento dei dati nel grafico in corso . . ."
  MDIFd.sbrStato.Refresh
  Dim recSerie As Recordset
  Dim strSQLSerie As String
  Dim recCandle As Recordset
  Dim strSQLCandle As String
  Dim g1 As String
  Dim m1 As String
  Dim a1 As String
  Dim g2 As String
  Dim m2 As String
  Dim a2 As String
  Dim dtmPrimaData As Date
  Dim dtmSecondaData As Date
  Dim blnErrData As Boolean
  Dim intAmpiezza As Integer
  Dim vntDynMSChart() As Variant
  Dim strCampo As String
  Dim blnCandle As Boolean
  blnCandle = False
  If cmbTipoGraf.Text <> strTipi(5) Then
    strCampo = cmbCampo.Text
    For i = 1 To UBound(strDynCampi)
      If strCampo = strDynCampi(i, 2) Then
        strSQLSerie = "SELECT Data, " & strDynCampi(i, 1) & " FROM Nomi INNER JOIN Prezzi ON " _
          & "Nomi.Codice = Prezzi.Codice WHERE Nome = '" & cmbStocks.Text & "'"
        Exit For
      End If
    Next
  Else
    blnCandle = True
    strSQLSerie = "SELECT Data, Apertura, Alto, Basso, Ultimo FROM Nomi INNER JOIN Prezzi ON " _
      & "Nomi.Codice = Prezzi.Codice WHERE Nome = '" & cmbStocks.Text & "'"
    strSQLCandle = "SELECT Max(Alto), Min(Basso) FROM Nomi INNER JOIN Prezzi ON " _
      & "Nomi.Codice = Prezzi.Codice WHERE Nome = '" & cmbStocks.Text & "'"
  End If
  g1 = cmbD1.Text
  a1 = cmbY1.Text
  For i = 1 To 12
    If cmbM1.Text = strMesi(i) Then
      m1 = i
      Exit For
    End If
  Next
  g2 = cmbD2.Text
  a2 = cmbY2.Text
  For i = 1 To 12
    If cmbM2.Text = strMesi(i) Then
      m2 = i
      Exit For
    End If
  Next
  blnErrData = False
  If Val(m1) = 2 And Val(g1) > 28 And Val(a1) Mod 4 <> 0 Then blnErrData = True
  If Val(m1) = 2 And Val(g1) > 29 And Val(a1) Mod 4 = 0 Then blnErrData = True
  If Val(m1) = 4 And Val(g1) > 30 Then blnErrData = True
  If Val(m1) = 6 And Val(g1) > 30 Then blnErrData = True
  If Val(m1) = 9 And Val(g1) > 30 Then blnErrData = True
  If Val(m1) = 11 And Val(g1) > 30 Then blnErrData = True
  If Val(m2) = 2 And Val(g2) > 28 And Val(a2) Mod 4 <> 0 Then blnErrData = True
  If Val(m2) = 2 And Val(g2) > 29 And Val(a2) Mod 4 = 0 Then blnErrData = True
  If Val(m2) = 4 And Val(g2) > 30 Then blnErrData = True
  If Val(m2) = 6 And Val(g2) > 30 Then blnErrData = True
  If Val(m2) = 9 And Val(g2) > 30 Then blnErrData = True
  If Val(m2) = 11 And Val(g2) > 30 Then blnErrData = True
  If blnErrData = True Then
    o = MsgBox("Controllare che le date inserite siano valide.", vbExclamation, "Financial data")
    MousePointer = vbDefault
    Exit Sub
  Else
    dtmPrimaData = m1 & "/" & g1 & "/" & a1
    dtmSecondaData = m2 & "/" & g2 & "/" & a2
    If blnCandle = False Then
      strSQLSerie = strSQLSerie & " AND (data BETWEEN #" & dtmPrimaData & "# AND #" & dtmSecondaData & "#) ORDER BY Data"
    Else
      dtmSecondaData = dtmSecondaData + 1
      strSQLSerie = strSQLSerie & " AND (data BETWEEN #" & dtmPrimaData & "# AND #" & dtmSecondaData & "#) ORDER BY Data"
      strSQLCandle = strSQLCandle & " AND data BETWEEN #" & dtmPrimaData & "# AND #" & dtmSecondaData & "#"
    End If
  End If
  Set recSerie = dbFinData.OpenRecordset(strSQLSerie, dbOpenSnapshot, dbReadOnly)
  If recSerie.RecordCount = 0 Then
    o = MsgBox("Il numero di valori disponibili è insufficiente per poter visualizzare il grafico.", vbInformation, "Financial Data")
    MousePointer = vbDefault
    MDIFd.sbrStato.SimpleText = "Impostare i parametri desiderati"
    MDIFd.sbrStato.Refresh
    Exit Sub
  End If
  If blnCandle = True Then
    Set recCandle = dbFinData.OpenRecordset(strSQLCandle, dbOpenSnapshot, dbReadOnly)
    Dim intNulli As Integer
    If recCandle(0) <> "" And recCandle(1) <> "" Then
      dblCandleUnit = (recCandle(0) - recCandle(1)) / 400
      dblCandleMin = recCandle(1)
    Else
      o = MsgBox("Grafico non disponibile.", vbInformation, "Financial Data")
      MousePointer = vbDefault
      MDIFd.sbrStato.SimpleText = "Impostare i parametri desiderati"
      MDIFd.sbrStato.Refresh
      Exit Sub
    End If
    recSerie.MoveFirst
    i = 0
    intAmpiezza = 0
    Do Until recSerie.EOF = True
      i = i + 1
      If recSerie(0) <> "" And recSerie(1) <> "" And recSerie(2) <> "" And recSerie(3) <> "" And recSerie(4) <> "" Then
        intAmpiezza = intAmpiezza + 1
      End If
      recSerie.MoveNext
    Loop
    If intAmpiezza <= 2 Then
      o = MsgBox("Il numero di valori disponibili è insufficiente per poter visualizzare il grafico.", vbInformation, "Financial Data")
      MousePointer = vbDefault
      MDIFd.sbrStato.SimpleText = "Impostare i parametri desiderati"
      MDIFd.sbrStato.Refresh
      Exit Sub
    End If
    If intAmpiezza >= 140 Then
      o = MsgBox("Il numero di valori è troppo elevato per riuscire a visualizzare il grafico. ", vbInformation, "Financial Data")
      MousePointer = vbDefault
      MDIFd.sbrStato.SimpleText = "Impostare i parametri desiderati"
      MDIFd.sbrStato.Refresh
      Exit Sub
    End If
    ReDim udtCandle(1 To intAmpiezza)
    recSerie.MoveFirst
    i = 0
    intNulli = 0
    Do Until recSerie.EOF = True
      If recSerie(0) <> "" And recSerie(1) <> "" And recSerie(2) <> "" And recSerie(3) <> "" And recSerie(4) <> "" Then
        i = i + 1
        udtCandle(i).Data = recSerie(0)
        udtCandle(i).Apertura = 430 - CInt((recSerie(1) - recCandle(1)) / dblCandleUnit)
        udtCandle(i).Alto = 430 - CInt((recSerie(2) - recCandle(1)) / dblCandleUnit)
        udtCandle(i).Basso = 430 - CInt((recSerie(3) - recCandle(1)) / dblCandleUnit)
        udtCandle(i).Ultimo = 430 - CInt((recSerie(4) - recCandle(1)) / dblCandleUnit)
        recSerie.MoveNext
      Else
        recSerie.MoveNext
      End If
    Loop
    intCandleLargh = Int((700 / intAmpiezza) * (3 / 5))
    intCandleDist = Int((700 / intAmpiezza) * (2 / 5))
    frmCandleStick.Show
  Else
    recSerie.MoveFirst
    i = 0
    intAmpiezza = 0
    Do Until recSerie.EOF = True
      i = i + 1
      If recSerie(0) <> "" And recSerie(1) <> "" Then
        intAmpiezza = intAmpiezza + 1
      End If
      recSerie.MoveNext
    Loop
    If intAmpiezza <= 2 Then
      o = MsgBox("Il numero di valori disponibili è insufficiente per poter visualizzare il grafico.", vbInformation, "Financial Data")
      MousePointer = vbDefault
      MDIFd.sbrStato.SimpleText = "Impostare i parametri desiderati"
      MDIFd.sbrStato.Refresh
      Exit Sub
    End If
    ReDim vntDynMSChart(1 To intAmpiezza, 1 To 2)
    i = 1
    recSerie.MoveFirst
    Do Until recSerie.EOF = True
      If recSerie(0) <> "" And recSerie(1) <> "" Then
        i = i + 1
        vntDynMSChart(i, 1) = Format(recSerie(0), "dd/mm/yy") & "        "
        vntDynMSChart(i, 2) = recSerie(1)
      End If
      recSerie.MoveNext
    Loop
    vntDynMSChart(1, 2) = "Serie"
    MDIFd.sbrStato.Visible = False
    frmSerieStoricheChart.mscSerie = vntDynMSChart
    frmSerieStoricheChart.mscSerie.TitleText = cmbStocks.Text & ": " & cmbCampo.Text
    If cmbTipoGraf.Text = strTipi(1) Then
      frmSerieStoricheChart.mscSerie.chartType = VtChChartType2dBar
    ElseIf cmbTipoGraf.Text = strTipi(2) Then
      frmSerieStoricheChart.mscSerie.chartType = VtChChartType3dBar
    ElseIf cmbTipoGraf.Text = strTipi(3) Then
      frmSerieStoricheChart.mscSerie.chartType = VtChChartType2dLine
    ElseIf cmbTipoGraf.Text = strTipi(4) Then
      frmSerieStoricheChart.mscSerie.chartType = VtChChartType3dLine
    End If
    frmSerieStoricheChart.Show
  End If
Subscript:
  Select Case Err.Number
    Case 9
      Resume Next
  End Select
End Sub
