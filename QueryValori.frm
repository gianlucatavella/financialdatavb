VERSION 5.00
Begin VB.Form frmQueryValori 
   Caption         =   "Impostazione query: valori dei campi"
   ClientHeight    =   7410
   ClientLeft      =   75
   ClientTop       =   405
   ClientWidth     =   11610
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   11610
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
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   6360
      Width           =   1935
   End
   Begin VB.CommandButton cmdEseguiQuery 
      Caption         =   "&Esegui query"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9600
      TabIndex        =   12
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtra i record in base ai valori"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   9735
      Begin VB.ComboBox cmbCriterio 
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
         Height          =   315
         Index           =   1
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2160
         Width           =   855
      End
      Begin VB.ComboBox cmbEO 
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
         Height          =   315
         Index           =   1
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   3000
         Width           =   735
      End
      Begin VB.TextBox txtValore 
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
         Height          =   315
         Index           =   1
         Left            =   7200
         TabIndex        =   7
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtValore 
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
         Height          =   315
         Index           =   2
         Left            =   7200
         TabIndex        =   11
         Top             =   3000
         Width           =   1695
      End
      Begin VB.ComboBox cmbCampo 
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
         Height          =   315
         Index           =   2
         Left            =   1800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3000
         Width           =   3615
      End
      Begin VB.ComboBox cmbCriterio 
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
         Height          =   315
         Index           =   2
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3000
         Width           =   855
      End
      Begin VB.ComboBox cmbCampo 
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
         Height          =   315
         Index           =   1
         Left            =   1800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2160
         Width           =   3615
      End
      Begin VB.ComboBox cmbEO 
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
         Height          =   315
         Index           =   0
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtValore 
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
         Index           =   0
         Left            =   7200
         TabIndex        =   3
         Top             =   1440
         Width           =   1695
      End
      Begin VB.ComboBox cmbCriterio 
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
         Index           =   0
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1440
         Width           =   855
      End
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
         Index           =   0
         Left            =   1800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "Valore"
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
         Left            =   7680
         TabIndex        =   17
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Criterio"
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
         Left            =   6000
         TabIndex        =   16
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Campo"
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
         Left            =   3240
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmQueryValori"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim j As Integer
Dim strCriterio(1 To 6) As String
Private Sub cmdAnnulla_Click()
  ScaricaFormsAttivi
End Sub
Private Sub cmdEseguiQuery_Click()
    Dim strCampiV(0 To 2) As String
    Dim strCampo As String
    Dim intPosComma As Integer
    Dim intLungComma As Integer
    strSQLvalori = ""
    Dim strEo As String
    For i = 0 To 2
      If cmbCampo(i).Text <> "" And cmbCriterio(i).Text <> "" And txtValore(i).Text <> "" And cmbCampo(i).Enabled = True Then
        intPosComma = InStr(1, txtValore(i).Text, ",")
        If intPosComma <> 0 Then
          intLungComma = Len(txtValore(i).Text)
          txtValore(i).Text = Left(txtValore(i).Text, intPosComma - 1) & "." & Right(txtValore(i).Text, intLungComma - intPosComma)
        End If
        If i > 0 Then strEo = " AND "
        If cmbEO(0) = "O" And i = 1 Then strEo = " OR "
        If cmbEO(1) = "O" And i = 2 Then strEo = " OR "
        For j = 1 To UBound(strDynCampi)
          strCampo = cmbCampo(i).Text
          If strCampo = strDynCampi(j, 2) Then
            strSQLvalori = strSQLvalori & strEo & strDynCampi(j, 1) & " " & cmbCriterio(i).Text & " " & Val(txtValore(i).Text)
          End If
        Next
      End If
    Next
    If strSQLstocks = "" And strSQLdata = "" And strSQLvalori <> "" Then
      strSQL = strSQL & " WHERE " & strSQLvalori
    Else
      If strSQLvalori = "" Then
        strSQL = strSQL & strSQLstocks & strSQLdata
      Else
        strSQL = strSQL & strSQLstocks & strSQLdata & " AND (" & strSQLvalori & ")"
      End If
    End If
    strSQL = strSQL & " ORDER BY Nome, Data "
  MousePointer = vbHourglass
  MDIFd.sbrStato.SimpleText = "Esecuzione della query in corso..."
  MDIFd.sbrStato.Refresh
  Set recQuery = dbFinData.OpenRecordset(strSQL, dbOpenSnapshot, dbReadOnly)
  MousePointer = 0
  If recQuery.RecordCount = 0 Then
    o = MsgBox("Nessun record soddisfa i parametri impostati nella query", vbInformation, "Financial Data")
  Else
    MousePointer = vbHourglass
    frmQueryRisultati.Show
  End If
  frmQueryValori.Hide
  frmQueryDate.Hide
  frmQueryCampi.Hide
End Sub
Private Sub cmdPrecedente_Click()
  frmQueryValori.Hide
End Sub
Private Sub Form_Activate()
  MousePointer = vbDefault
  udtFormsLoad.Qvalori = True
  MDIFd.sbrStato.Visible = True
  MDIFd.sbrStato.SimpleText = "Per eseguire premere Esegui query"
  MDIFd.sbrStato.Refresh
End Sub
Private Sub Form_Load()
  udtFormsLoad.Qvalori = True
  MousePointer = vbDefault
  strCriterio(1) = "="
  strCriterio(2) = ">"
  strCriterio(3) = ">="
  strCriterio(4) = "<="
  strCriterio(5) = "<"
  strCriterio(6) = "<>"
  For i = 0 To 2
    cmbCampo(i).AddItem ""
    cmbCriterio(i).AddItem ""
  Next
  For i = 0 To 1
    cmbEO(i).AddItem "E"
    cmbEO(i).AddItem "O"
  Next
  For i = 0 To 2
    For j = 1 To UBound(strDynCampi)
      cmbCampo(i).AddItem strDynCampi(j, 2)
    Next
  Next
  For i = 0 To 2
    For j = 1 To 6
      cmbCriterio(i).AddItem strCriterio(j)
    Next
  Next
End Sub
Private Sub txtValore_Change(Index As Integer)
  If Index < 2 Then
    If txtValore(Index).Text <> "" And cmbCampo(Index).Text <> "" And cmbCriterio(Index).Text <> "" Then
      txtValore(Index + 1).Enabled = True
      cmbCampo(Index + 1).Enabled = True
      cmbCriterio(Index + 1).Enabled = True
      cmbEO(Index).Enabled = True
        If Index = 0 And txtValore(Index + 1).Text <> "" And cmbCampo(Index + 1).Text <> "" And cmbCriterio(Index + 1).Text <> "" Then
          txtValore(Index + 2).Enabled = True
          cmbCampo(Index + 2).Enabled = True
          cmbCriterio(Index + 2).Enabled = True
          cmbEO(Index + 1).Enabled = True
        End If
    Else
      If Index = 1 Then
        txtValore(Index + 1).Enabled = False
        cmbCampo(Index + 1).Enabled = False
        cmbCriterio(Index + 1).Enabled = False
        cmbEO(Index).Enabled = False
      Else
        For i = 1 To 2
          txtValore(Index + i).Enabled = False
          cmbCampo(Index + i).Enabled = False
          cmbCriterio(Index + i).Enabled = False
          cmbEO(Index + i - 1).Enabled = False
        Next
      End If
    End If
  End If
End Sub
Private Sub cmbCampo_Click(Index As Integer)
  If Index < 2 Then
    If txtValore(Index).Text <> "" And cmbCampo(Index).Text <> "" And cmbCriterio(Index).Text <> "" Then
      txtValore(Index + 1).Enabled = True
      cmbCampo(Index + 1).Enabled = True
      cmbCriterio(Index + 1).Enabled = True
      cmbEO(Index).Enabled = True
      If Index = 0 And txtValore(Index + 1).Text <> "" And cmbCampo(Index + 1).Text <> "" And cmbCriterio(Index + 1).Text <> "" Then
          txtValore(Index + 2).Enabled = True
          cmbCampo(Index + 2).Enabled = True
          cmbCriterio(Index + 2).Enabled = True
          cmbEO(Index + 1).Enabled = True
      End If
    Else
      If Index = 1 Then
        txtValore(Index + 1).Enabled = False
        cmbCampo(Index + 1).Enabled = False
        cmbCriterio(Index + 1).Enabled = False
        cmbEO(Index).Enabled = False
      Else
        For i = 1 To 2
          txtValore(Index + i).Enabled = False
          cmbCampo(Index + i).Enabled = False
          cmbCriterio(Index + i).Enabled = False
          cmbEO(Index + i - 1).Enabled = False
        Next
      End If
    End If
  End If
End Sub
Private Sub cmbCriterio_Click(Index As Integer)
  If Index < 2 Then
    If txtValore(Index).Text <> "" And cmbCampo(Index).Text <> "" And cmbCriterio(Index).Text <> "" Then
      txtValore(Index + 1).Enabled = True
      cmbCampo(Index + 1).Enabled = True
      cmbCriterio(Index + 1).Enabled = True
      cmbEO(Index).Enabled = True
      If Index = 0 And txtValore(Index + 1).Text <> "" And cmbCampo(Index + 1).Text <> "" And cmbCriterio(Index + 1).Text <> "" Then
          txtValore(Index + 2).Enabled = True
          cmbCampo(Index + 2).Enabled = True
          cmbCriterio(Index + 2).Enabled = True
          cmbEO(Index + 1).Enabled = True
      End If
    Else
      If Index = 1 Then
        txtValore(Index + 1).Enabled = False
        cmbCampo(Index + 1).Enabled = False
        cmbCriterio(Index + 1).Enabled = False
        cmbEO(Index).Enabled = False
      Else
        For i = 1 To 2
          txtValore(Index + i).Enabled = False
          cmbCampo(Index + i).Enabled = False
          cmbCriterio(Index + i).Enabled = False
          cmbEO(Index + i - 1).Enabled = False
        Next
      End If
    End If
  End If
End Sub
