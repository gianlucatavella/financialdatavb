VERSION 5.00
Begin VB.Form frmQueryDate 
   Caption         =   "Impostazione query: date"
   ClientHeight    =   7425
   ClientLeft      =   2805
   ClientTop       =   2070
   ClientWidth     =   11880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7425
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
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
      TabIndex        =   12
      Top             =   6360
      Width           =   1935
   End
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
      TabIndex        =   11
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
      TabIndex        =   1
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtra i record in base alle date"
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
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   11535
      Begin VB.ComboBox cmbIaa2 
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
         Left            =   10200
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2640
         Width           =   855
      End
      Begin VB.ComboBox cmbIaa1 
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
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2640
         Width           =   855
      End
      Begin VB.ComboBox cmbImm2 
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
         Left            =   8640
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2640
         Width           =   1455
      End
      Begin VB.ComboBox cmbImm1 
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
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2640
         Width           =   1335
      End
      Begin VB.ComboBox cmbIgg2 
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
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2640
         Width           =   735
      End
      Begin VB.ComboBox cmbIgg1 
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
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2640
         Width           =   735
      End
      Begin VB.ComboBox cmbSaa 
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
         Left            =   6240
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.ComboBox cmbSmm 
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
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1200
         Width           =   1335
      End
      Begin VB.ComboBox cmbSgg 
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
         Left            =   3960
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   735
      End
      Begin VB.OptionButton optQualsiasi 
         Caption         =   "Qualsiasi data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   15
         Top             =   3960
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton optIntervallo 
         Caption         =   "Intervallo temporale"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   14
         Top             =   2520
         Width           =   2295
      End
      Begin VB.OptionButton optSingola 
         Caption         =   "Singola data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   5
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "A"
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
         Left            =   7440
         TabIndex        =   17
         Top             =   2760
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
         Left            =   3480
         TabIndex        =   16
         Top             =   2760
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmQueryDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strMesi(1 To 12) As String
Private Sub cmbIaa1_Click()
  optIntervallo.Value = True
End Sub
Private Sub cmbIaa2_Click()
  optIntervallo.Value = True
End Sub
Private Sub cmbIgg1_Click()
  optIntervallo.Value = True
End Sub
Private Sub cmbIgg2_Click()
  optIntervallo.Value = True
End Sub
Private Sub cmbImm1_Click()
  optIntervallo.Value = True
End Sub
Private Sub cmbImm2_Click()
  optIntervallo.Value = True
End Sub
Private Sub cmbSaa_Click()
  optSingola.Value = True
End Sub
Private Sub cmbSgg_Click()
  optSingola.Value = True
End Sub
Private Sub cmbSmm_Click()
  optSingola.Value = True
End Sub
Private Sub cmdAnnulla_Click()
  ScaricaFormsAttivi
End Sub
Private Sub cmdPrecedente_Click()
  frmQueryDate.Hide
End Sub
Private Sub cmdSuccessivo_Click()
  Dim g As String
  Dim m As String
  Dim a As String
  Dim g1 As String
  Dim m1 As String
  Dim a1 As String
  Dim g2 As String
  Dim m2 As String
  Dim a2 As String
  Dim dtmDataSingola As Date
  Dim dtmPrimaData As Date
  Dim dtmSecondaData As Date
  Dim blnErrData As Boolean
  If optQualsiasi.Value = True Then
    strSQLdata = ""
    frmQueryValori.Show
  End If
  strSQLdata = ""
  If optSingola.Value = True Then
    g = cmbSgg.Text
    a = cmbSaa.Text
    For i = 1 To 12
      If cmbSmm.Text = strMesi(i) Then
        m = i
        Exit For
      End If
    Next
    If g <> "" And m <> "" And a <> "" Then
      blnErrData = False
      If Val(m) = 2 And Val(g) > 28 And Val(a) Mod 4 <> 0 Then blnErrData = True
      If Val(m) = 2 And Val(g) > 29 And Val(a) Mod 4 = 0 Then blnErrData = True
      If Val(m) = 4 And Val(g) > 30 Then blnErrData = True
      If Val(m) = 6 And Val(g) > 30 Then blnErrData = True
      If Val(m) = 9 And Val(g) > 30 Then blnErrData = True
      If Val(m) = 11 And Val(g) > 30 Then blnErrData = True
      If blnErrData = True Then
        o = MsgBox("Controllare che la data inserita sia valida.", vbExclamation, "Financial data")
      Else
        dtmDataSingola = m & "/" & g & "/" & a
        If strSQLstocks = "" Then
          strSQLdata = " WHERE (data = #" & dtmDataSingola & "#) "
        Else
          strSQLdata = " AND (data = #" & dtmDataSingola & "#) "
        End If
        frmQueryValori.Show
      End If
    Else
      o = MsgBox("La data non è stata completamente inserita.", vbInformation, "Financial data")
    End If
  End If
  If optIntervallo.Value = True Then
    g1 = cmbIgg1.Text
    a1 = cmbIaa1.Text
    For i = 1 To 12
      If cmbImm1.Text = strMesi(i) Then
        m1 = i
        Exit For
      End If
    Next
    g2 = cmbIgg2.Text
    a2 = cmbIaa2.Text
    For i = 1 To 12
      If cmbImm2.Text = strMesi(i) Then
        m2 = i
        Exit For
      End If
    Next
    If g1 <> "" And m1 <> "" And a1 <> "" And g2 <> "" And m2 <> "" And a2 <> "" Then
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
      Else
        dtmPrimaData = m1 & "/" & g1 & "/" & a1
        dtmSecondaData = m2 & "/" & g2 & "/" & a2
        If strSQLstocks = "" Then
          strSQLdata = " WHERE (data BETWEEN #" & dtmPrimaData & "# AND #" & dtmSecondaData & "#) "
        Else
          strSQLdata = " AND (data BETWEEN #" & dtmPrimaData & "# AND #" & dtmSecondaData & "#) "
        End If
        frmQueryValori.Show
      End If
    Else
      o = MsgBox("Le due date non sono state completamente inserite.", vbInformation, "Financial data")
    End If
  End If
End Sub
Private Sub Form_Activate()
  udtFormsLoad.Qdate = True
  MDIFd.sbrStato.Visible = True
  MDIFd.sbrStato.SimpleText = "Selezionare l'opzione desiderata"
  MDIFd.sbrStato.Refresh
End Sub
Private Sub Form_Load()
  udtFormsLoad.Qdate = True
  Dim i As Integer
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
  For i = 1 To 12
    cmbSmm.AddItem strMesi(i)
    cmbImm1.AddItem strMesi(i)
    cmbImm2.AddItem strMesi(i)
  Next
  For i = 1 To 31
    cmbSgg.AddItem Str(i)
    cmbIgg1.AddItem Str(i)
    cmbIgg2.AddItem Str(i)
  Next
  For i = intMin To intMax
    cmbSaa.AddItem Str(i)
    cmbIaa1.AddItem Str(i)
    cmbIaa2.AddItem Str(i)
  Next
End Sub

