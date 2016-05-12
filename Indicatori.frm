VERSION 5.00
Begin VB.Form frmIndicatori 
   Caption         =   "Indicatori descrittivi tipici del mercato azionario: impostazione parametri"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   11610
   WindowState     =   2  'Maximized
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
      TabIndex        =   3
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Applicazione di computazioni specifiche ai record ottenuti"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   480
      TabIndex        =   0
      Top             =   600
      Width           =   10455
      Begin VB.ComboBox cmbRSICampo 
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
         TabIndex        =   14
         Top             =   1320
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.ComboBox cmbMMGiorni 
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
         Left            =   8400
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbMMCampo 
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
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.CommandButton cmdAnnulla 
         Caption         =   "&Annulla"
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
         Left            =   5760
         TabIndex        =   2
         Top             =   4080
         Width           =   1935
      End
      Begin VB.CommandButton cmdCalcola 
         Caption         =   "Ca&lcola"
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
         Left            =   8160
         TabIndex        =   1
         Top             =   4080
         Width           =   1935
      End
      Begin VB.ComboBox cmbRocSett 
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
         Left            =   8400
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cmbRocCampo 
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
         TabIndex        =   6
         Top             =   1320
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.ComboBox cmbIndicatori 
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
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label lblRSICampo 
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
         Height          =   255
         Left            =   5160
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblMMCampo 
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
         Height          =   255
         Left            =   5160
         TabIndex        =   13
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblMMGiorni 
         Caption         =   "Intervallo (giorni)"
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
         Left            =   7920
         TabIndex        =   12
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblRocSett 
         Caption         =   "Intervallo (settimane)"
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
         Left            =   7920
         TabIndex        =   9
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label lblRocCampo 
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
         Height          =   255
         Left            =   5160
         TabIndex        =   7
         Top             =   960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label lblIndicatore 
         Caption         =   "Indicatori disponibili"
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
         Left            =   1560
         TabIndex        =   4
         Top             =   960
         Width           =   2295
      End
   End
End
Attribute VB_Name = "frmIndicatori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strIndicatori(1 To 4) As String
Dim blnAnnulla As Boolean
Dim i As Long
Private Sub cmbIndicatori_Click()
  If cmbIndicatori.Text = strIndicatori(2) Then
    cmdCalcola.Enabled = True
  Else
    cmdCalcola.Enabled = False
  End If
  If cmbIndicatori.Text = strIndicatori(3) Then
    lblMMGiorni.Visible = True
    lblMMCampo.Visible = True
    cmbMMCampo.Visible = True
    cmbMMGiorni.Visible = True
    If cmbMMCampo.Text <> "" And cmbMMGiorni.Text <> "" Then
      cmdCalcola.Enabled = True
    End If
  Else
    lblMMGiorni.Visible = False
    lblMMCampo.Visible = False
    cmbMMCampo.Visible = False
    cmbMMGiorni.Visible = False
  End If
  If cmbIndicatori.Text = strIndicatori(1) Then
    lblRocSett.Visible = True
    lblRocCampo.Visible = True
    cmbRocCampo.Visible = True
    cmbRocSett.Visible = True
    If cmbRocCampo.Text <> "" And cmbRocSett.Text <> "" Then
      cmdCalcola.Enabled = True
    End If
  Else
    lblRocSett.Visible = False
    lblRocCampo.Visible = False
    cmbRocCampo.Visible = False
    cmbRocSett.Visible = False
  End If
  If cmbIndicatori.Text = strIndicatori(4) Then
    lblRSICampo.Visible = True
    cmbRSICampo.Visible = True
    If cmbRSICampo.Text <> "" Then
      cmdCalcola.Enabled = True
    End If
  Else
    lblRSICampo.Visible = False
    cmbRSICampo.Visible = False
  End If
End Sub
Private Sub cmbMMCampo_Click()
  If cmbMMCampo.Text <> "" And cmbMMGiorni <> "" Then
    cmdCalcola.Enabled = True
  Else
    cmdCalcola.Enabled = False
  End If
End Sub
Private Sub cmbMMGiorni_Click()
  If cmbMMCampo.Text <> "" And cmbMMGiorni <> "" Then
    cmdCalcola.Enabled = True
  Else
    cmdCalcola.Enabled = False
  End If
End Sub
Private Sub cmbRocCampo_Click()
  If cmbRocCampo.Text <> "" And cmbRocSett <> "" Then
    cmdCalcola.Enabled = True
  Else
    cmdCalcola.Enabled = False
  End If
End Sub
Private Sub cmbRocSett_Click()
  If cmbRocCampo.Text <> "" And cmbRocSett <> "" Then
    cmdCalcola.Enabled = True
  Else
    cmdCalcola.Enabled = False
  End If
End Sub
Private Sub cmbRSICampo_Click()
  If cmbRSICampo.Text <> "" Then
    cmdCalcola.Enabled = True
  Else
    cmdCalcola.Enabled = False
  End If
End Sub
Private Sub cmdAnnulla_Click()
  blnAnnulla = True
End Sub
Private Sub cmdCalcola_Click()
  MDIFd.sbrStato.Visible = True
  MDIFd.sbrStato.SimpleText = "Computazione dell'indicatore in corso . . ."
  MDIFd.sbrStato.Refresh
  MousePointer = 13
  If cmbIndicatori.Text = strIndicatori(1) Then
    cmdAnnulla.Enabled = True
    cmdCalcola.Enabled = False
    CalcolaROC
    cmdCalcola.Enabled = True
    cmdAnnulla.Enabled = False
  End If
  If cmbIndicatori.Text = strIndicatori(2) Then
    cmdAnnulla.Enabled = True
    cmdCalcola.Enabled = False
    CalcolaStocastico
    cmdCalcola.Enabled = True
    cmdAnnulla.Enabled = False
  End If
  If cmbIndicatori.Text = strIndicatori(3) Then
    cmdAnnulla.Enabled = True
    cmdCalcola.Enabled = False
    CalcolaMediaMobile
    cmdCalcola.Enabled = True
    cmdAnnulla.Enabled = False
  End If
  If cmbIndicatori.Text = strIndicatori(4) Then
    cmdAnnulla.Enabled = True
    cmdCalcola.Enabled = False
    CalcolaRSI
    cmdCalcola.Enabled = True
    cmdAnnulla.Enabled = False
  End If
  MDIFd.sbrStato.Visible = False
  MousePointer = 0
End Sub
Private Sub cmdChiudi_Click()
  udtFormsLoad.Icalcola = False
  frmIndicatori.Hide
  Unload frmIndicatori
End Sub
Private Sub Form_Activate()
  udtFormsLoad.Icalcola = True
  MDIFd.mnuGraficoIndicatori.Enabled = False
  MDIFd.tlbBarraStrumenti.Buttons(4).Enabled = False
End Sub
Private Sub Form_Load()
  udtFormsLoad.Icalcola = True
  MDIFd.mnuGraficoIndicatori.Enabled = False
  MDIFd.tlbBarraStrumenti.Buttons(4).Enabled = False
  strIndicatori(1) = "Rate of Change"
  strIndicatori(2) = "Indice Stocastico"
  strIndicatori(3) = "Media Mobile"
  strIndicatori(4) = "Relative Strenght Index"
  For i = 1 To 4
    cmbIndicatori.AddItem strIndicatori(i)
  Next
  For i = 1 To UBound(udtDynCampi)
    cmbRocCampo.AddItem udtDynCampi(i).Descrizione
    cmbMMCampo.AddItem udtDynCampi(i).Descrizione
    If udtDynCampi(i).Descrizione <> "Volume accumulato" Then
      cmbRSICampo.AddItem udtDynCampi(i).Descrizione
    End If
  Next
  For i = 1 To 52
    cmbRocSett.AddItem i
  Next
  For i = 5 To 120 Step 5
    cmbMMGiorni.AddItem i
  Next
End Sub
Private Sub CalcolaROC()
  Dim dtmDataROC As Date
  Dim recROC As Recordset
  Dim strSQLindROC As String
  Dim intCampo As Integer
  blnAnnulla = False
  MDIFd.pbrAvanz.Visible = True
  ReDim dblDynIndicatore(0 To recQuery.RecordCount - 1)
  recQuery.MoveFirst
  intAmpGrafico = 1
  For i = 1 To UBound(udtDynCampi)
    If cmbRocCampo.Text = udtDynCampi(i).Descrizione Then
      strCampo = udtDynCampi(i).Nome
      intCampo = udtDynCampi(i).Field
      Exit For
    End If
  Next
  For i = 0 To recQuery.RecordCount - 1
    If i Mod 200 = 0 Then DoEvents
    If blnAnnulla = True Then
      MDIFd.pbrAvanz.Visible = False
      Exit Sub
    End If
    dtmDataROC = recQuery(1) - (Val(cmbRocSett.Text) * 7)
    strSQLindROC = "SELECT " & strCampo & " FROM Nomi INNER JOIN Prezzi ON Nomi.codice = Prezzi.codice WHERE "
    strSQLindROC = strSQLindROC & " nome = '" & recQuery(0) & "' AND data = #" & dtmDataROC & "# "
    Set recROC = dbFinData.OpenRecordset(strSQLindROC, dbOpenSnapshot, dbReadOnly)
    If recROC.RecordCount <> 0 Then
      If recROC(0) <> "" And recQuery(intCampo) <> "" And recROC(0) <> 0 Then
        dblDynIndicatore(i) = (Val(recQuery(intCampo)) / Val(recROC(0))) * 100
        intAmpGrafico = intAmpGrafico + 1
      Else
        dblDynIndicatore(i) = -1
      End If
    Else
      dblDynIndicatore(i) = -1
    End If
    MDIFd.pbrAvanz.Value = recQuery.PercentPosition
    recQuery.MoveNext
  Next
  MDIFd.pbrAvanz.Visible = False
  frmIndicatoriRis.Show
End Sub
Private Sub CalcolaMediaMobile()
  Dim dtmDataMM As Date
  Dim recMM As Recordset
  Dim strSQLindMM As String
  Dim intCampo As Integer
  blnAnnulla = False
  MDIFd.pbrAvanz.Visible = True
  ReDim dblDynIndicatore(0 To recQuery.RecordCount - 1)
  recQuery.MoveFirst
  intAmpGrafico = 1
  For i = 1 To UBound(udtDynCampi)
    If cmbMMCampo.Text = udtDynCampi(i).Descrizione Then
      strCampo = udtDynCampi(i).Nome
      intCampo = udtDynCampi(i).Field
      Exit For
    End If
  Next
  For i = 0 To recQuery.RecordCount - 1
    If i Mod 200 = 0 Then DoEvents
    If blnAnnulla = True Then
      MDIFd.pbrAvanz.Visible = False
      Exit Sub
    End If
    dtmDataMM = recQuery(1) - Val(cmbMMGiorni.Text)
    strSQLindMM = "SELECT Avg(" & strCampo & ") FROM Nomi INNER JOIN Prezzi ON Nomi.codice = Prezzi.codice WHERE "
    strSQLindMM = strSQLindMM & " nome = '" & recQuery(0) & "' AND data BETWEEN #" & dtmDataMM & "# AND #" & recQuery(1) & "# "
    Set recMM = dbFinData.OpenRecordset(strSQLindMM, dbOpenSnapshot, dbReadOnly)
    If recMM.RecordCount > 0 Then
      If recMM(0) <> "" And recQuery(intCampo) <> "" And recMM(0) <> 0 Then
        dblDynIndicatore(i) = recMM(0)
        intAmpGrafico = intAmpGrafico + 1
      Else
        dblDynIndicatore(i) = -1
      End If
    Else
      dblDynIndicatore(i) = -1
    End If
    MDIFd.pbrAvanz.Value = recQuery.PercentPosition
    recQuery.MoveNext
  Next
  MDIFd.pbrAvanz.Visible = False
  frmIndicatoriRis.Show
End Sub
Private Sub CalcolaStocastico()
  Dim blnFieldUltimo As Boolean
  blnFieldUltimo = False
  For i = 1 To UBound(udtDynCampi)
    If udtDynCampi(i).Nome = "Ultimo" Then
      blnFieldUltimo = True
    End If
  Next
  If blnFieldUltimo = False Then
    o = MsgBox("Per procedere al calcolo dell'Indice Stocastico è necessario includere nei campi da visualizzare" _
    & " il campo denominato 'Ultimo prezzo'.", vbInformation, "Financial Data")
    Exit Sub
  End If
  Dim recStocastico As Recordset
  Dim intCampo As Integer
  Dim intPrimK As Integer
  Dim strSQLindStocastico As String
  Dim lngPK As Long
  blnAnnulla = False
  MDIFd.pbrAvanz.Visible = True
  ReDim dblDynStocastico(0 To recQuery.RecordCount - 1)
  recQuery.MoveFirst
  intAmpGrafico = 1
  For i = 1 To UBound(udtDynCampi)
    If udtDynCampi(i).Nome = "Ultimo" Then
      intCampo = udtDynCampi(i).Field
      Exit For
    End If
  Next
  intPrimK = udtDynCampi(UBound(udtDynCampi)).Field + 1
  For i = 0 To recQuery.RecordCount - 1
    If i Mod 200 = 0 Then DoEvents
    If blnAnnulla = True Then
      MDIFd.pbrAvanz.Visible = False
      Exit Sub
    End If
    lngPK = recQuery(intPrimK) - 4
    strSQLindStocastico = "SELECT Min(Basso), Max(Alto), Count(Alto) FROM Nomi INNER JOIN Prezzi ON Nomi.codice = Prezzi.codice WHERE" _
    & " nome = '" & recQuery(0) & "' AND PrimaryK BETWEEN " & recQuery(intPrimK) & " AND " _
    & lngPK
    Set recStocastico = dbFinData.OpenRecordset(strSQLindStocastico, dbOpenSnapshot, dbReadOnly)
    If recStocastico.RecordCount > 0 And recStocastico(2) = 5 Then
      dblDynStocastico(i).CL5 = recQuery(intCampo) - recStocastico(0)
      dblDynStocastico(i).H5L5 = recStocastico(1) - recStocastico(0)
      dblDynStocastico(i).K = 100 * (dblDynStocastico(i).CL5 / dblDynStocastico(i).H5L5)
    Else
      dblDynStocastico(i).K = -1
    End If
    MDIFd.pbrAvanz.Value = recQuery.PercentPosition
    recQuery.MoveNext
  Next
  For i = 3 To UBound(dblDynStocastico)
    If dblDynStocastico(i).K <> -1 Then
      dblDynStocastico(i).D = 100 * ((dblDynStocastico(i).CL5 + dblDynStocastico(i - 1).CL5 + dblDynStocastico(i - 2).CL5) / (dblDynStocastico(i).H5L5 + dblDynStocastico(i - 1).H5L5 + dblDynStocastico(i - 2).H5L5))
    End If
  Next
  MDIFd.pbrAvanz.Visible = False
  frmIndicatoriRis.Show
End Sub
Private Sub CalcolaRSI()
  Dim j As Integer
  Dim recRSI As Recordset
  Dim strSQLindRSI As String
  Dim intCampo As Integer
  Dim dtmDataRSI As Date
  Dim dblDynRSI() As Double
  Dim dblIncrementi As Double
  Dim dblDecrementi As Double
  Dim intDimin As Integer
  Dim blnRecVuoto As Boolean
  Dim intIncrementi As Integer
  Dim intDecrementi As Integer
  blnAnnulla = False
  MDIFd.pbrAvanz.Visible = True
  ReDim dblDynIndicatore(0 To recQuery.RecordCount - 1)
  recQuery.MoveFirst
  intAmpGrafico = 1
  For i = 1 To UBound(udtDynCampi)
    If cmbRSICampo.Text = udtDynCampi(i).Descrizione Then
      strCampo = udtDynCampi(i).Nome
      intCampo = udtDynCampi(i).Field
      Exit For
    End If
  Next
  For i = 0 To recQuery.RecordCount - 1
    If i Mod 200 = 0 Then DoEvents
    If blnAnnulla = True Then
      MDIFd.pbrAvanz.Visible = False
      Exit Sub
    End If
    dtmDataRSI = recQuery(1) - 14
    strSQLindRSI = "SELECT " & strCampo & " FROM Nomi INNER JOIN Prezzi ON Nomi.codice = Prezzi.codice WHERE " _
    & " nome = '" & recQuery(0) & "' AND data BETWEEN #" & dtmDataRSI & "# AND #" & recQuery(1) & "# ORDER BY PrimaryK"
    Set recRSI = dbFinData.OpenRecordset(strSQLindRSI, dbOpenSnapshot, dbReadOnly)
    If recRSI.RecordCount > 1 Then
      recRSI.MoveFirst
      ReDim dblDynRSI(1 To 2, 1 To recRSI.RecordCount)
      j = 0
      intDimin = 0
      blnRecVuoto = False
      Do Until recRSI.EOF = True
        j = j + 1
        If recRSI(0) <> "" Then
          dblDynRSI(1, j) = recRSI(0)
        ElseIf intDimin > recRSI.RecordCount + 1 Then
          intDimin = intDimin + 1
          ReDim Preserve dblDynRSI(1 To 2, 1 To recRSI.RecordCount - intDimin)
          j = j - 1
        Else
          blnRecVuoto = True
        End If
        recRSI.MoveNext
      Loop
      dblIncrementi = 0
      dblDecrementi = 0
      intIncrementi = 0
      intDecrementi = 0
      For j = 1 To UBound(dblDynRSI) - 1
        dblDynRSI(2, j) = dblDynRSI(1, j + 1) - dblDynRSI(1, j)
        If dblDynRSI(2, j) > 0 Then
          intIncrementi = intIncrementi + 1
          dblIncrementi = dblIncrementi + dblDynRSI(2, j)
        Else
          intDecrementi = intDecrementi + 1
          dblDecrementi = dblDecrementi + Abs(dblDynRSI(2, j))
        End If
      Next
      If intIncrementi = 0 Then
        intIncrementi = 1
      End If
      If intDecrementi = 0 Then
        intDecrementi = 1
      End If
      If blnRecVuoto = False Then
        dblDynIndicatore(i) = 100 - (100 / (1 + ((Exp(dblIncrementi) / intIncrementi) / (Exp(dblDecrementi) / intDecrementi))))
      Else
        dblDynIndicatore(i) = -1
      End If
    Else
      dblDynIndicatore(i) = -1
    End If
    MDIFd.pbrAvanz.Value = recQuery.PercentPosition
    recQuery.MoveNext
  Next
  MDIFd.pbrAvanz.Visible = False
  frmIndicatoriRis.Show
End Sub
