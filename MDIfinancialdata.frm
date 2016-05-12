VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIFd 
   BackColor       =   &H8000000C&
   Caption         =   "Financial Data"
   ClientHeight    =   5610
   ClientLeft      =   2775
   ClientTop       =   1710
   ClientWidth     =   7365
   Icon            =   "MDIfinancialdata.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar tlbBarraStrumenti 
      Align           =   1  'Align Top
      DragIcon        =   "MDIfinancialdata.frx":0442
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   1164
      ButtonWidth     =   1244
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "imgIconeBarra"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ApriConnessione"
            Object.ToolTipText     =   "Connessione al database"
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ImpostaQuery"
            Object.ToolTipText     =   "Nuova query"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "ModificaQuery"
            Object.ToolTipText     =   "Modifica query"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "GraficoIndicatore"
            Object.ToolTipText     =   "Indicatore "
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "GraficoSerieStorica"
            Object.ToolTipText     =   "Serie storica "
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
      OLEDropMode     =   1
      Begin MSComDlg.CommonDialog dlgConnessione 
         Left            =   3600
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin ComctlLib.ProgressBar pbrAvanz 
         Height          =   255
         Left            =   6120
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   450
         _Version        =   327682
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin ComctlLib.ImageList imgIconeBarra 
         Left            =   2040
         Top             =   1680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   10
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIfinancialdata.frx":0884
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIfinancialdata.frx":0B9E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIfinancialdata.frx":0EB8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIfinancialdata.frx":11D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIfinancialdata.frx":14EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIfinancialdata.frx":1806
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIfinancialdata.frx":1B20
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIfinancialdata.frx":1E3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIfinancialdata.frx":2154
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "MDIfinancialdata.frx":246E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin ComctlLib.StatusBar sbrStato 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5355
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuDatabase 
      Caption         =   "&Database"
      Begin VB.Menu mnuDatabaseApri 
         Caption         =   "&Apri connessione"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuDatabaseChiudi 
         Caption         =   "&Chiudi connessione"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
         Index           =   12
      End
      Begin VB.Menu mnuDatabaseEsci 
         Caption         =   "&Esci"
      End
   End
   Begin VB.Menu mnuQuery 
      Caption         =   "&Query"
      Begin VB.Menu mnuQueryNuova 
         Caption         =   "&Nuova"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuQueryModifica 
         Caption         =   "&Modifica"
      End
   End
   Begin VB.Menu mnuGrafico 
      Caption         =   "Gra&fico"
      Begin VB.Menu mnuGraficoSerie 
         Caption         =   "&Serie storica"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuGraficoIndicatori 
         Caption         =   "&Indicatore"
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGraficoChiudi 
         Caption         =   "&Chiudi"
      End
   End
   Begin VB.Menu mnuGuida 
      Caption         =   "&Guida"
      Begin VB.Menu mnuGuidaArgomenti 
         Caption         =   "&Argomenti"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuGuidaManuale 
         Caption         =   "&Manuale"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGuidaInfo 
         Caption         =   "&Info su Financial Data"
      End
   End
End
Attribute VB_Name = "MDIFd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
  udtFormsLoad.Icalcola = False
  udtFormsLoad.Ichart = False
  udtFormsLoad.Irisultati = False
  udtFormsLoad.Qcampi = False
  udtFormsLoad.Qdate = False
  udtFormsLoad.Qrisultati = False
  udtFormsLoad.Qstocks = False
  udtFormsLoad.Qvalori = False
  udtFormsLoad.Sstoriche = False
  sbrStato.SimpleText = "Financial Data 1.0"
  sbrStato.Refresh
  mnuDatabaseChiudi.Enabled = False
  mnuQuery.Enabled = False
  mnuQueryModifica.Enabled = False
  mnuGrafico.Enabled = False
  mnuGraficoIndicatori.Enabled = False
  mnuGraficoChiudi.Enabled = False
  tlbBarraStrumenti.Buttons(2).Enabled = False
  tlbBarraStrumenti.Buttons(3).Enabled = False
  tlbBarraStrumenti.Buttons(4).Enabled = False
  tlbBarraStrumenti.Buttons(5).Enabled = False
End Sub
Private Sub mnuDatabaseApri_Click()
  ConnessioneDB
End Sub
Private Sub mnuDatabaseChiudi_Click()
  ScaricaFormsAttivi
  dbFinData.Close
  sbrStato.SimpleText = "Financial Data 1.0"
  sbrStato.Refresh
  mnuDatabaseChiudi.Enabled = False
  mnuQuery.Enabled = False
  mnuQueryModifica.Enabled = False
  mnuGrafico.Enabled = False
  mnuGraficoIndicatori.Enabled = False
  mnuGraficoChiudi.Enabled = False
  tlbBarraStrumenti.Buttons(1).Enabled = True
  tlbBarraStrumenti.Buttons(2).Enabled = False
  tlbBarraStrumenti.Buttons(3).Enabled = False
  tlbBarraStrumenti.Buttons(4).Enabled = False
  tlbBarraStrumenti.Buttons(5).Enabled = False
  mnuDatabaseApri.Enabled = True
End Sub
Private Sub mnuGraficoChiudi_Click()
  mnuGraficoChiudi.Enabled = False
  If udtFormsLoad.CandleStick = True Then
    frmCandleStick.Hide
  End If
  If udtFormsLoad.Ichart = True Then
    frmIndicatoriChart.Hide
  End If
  If udtFormsLoad.SstoricheChart = True Then
    frmSerieStoricheChart.Hide
 End If
End Sub
Private Sub mnuGraficoIndicatori_Click()
  frmIndicatori.Show
End Sub
Private Sub mnuGraficoSerie_Click()
  GraficoSerie
End Sub
Private Sub mnuDatabaseEsci_Click()
  frmEsci.Show vbModal
End Sub
Private Sub mnuGuidaArgomenti_Click()
  frmGuida.Show
End Sub
Private Sub mnuGuidaManuale_Click()
  frmManuale.Show
End Sub
Private Sub mnuQueryNuova_Click()
  NuovaQuery
End Sub
Private Sub mnuQueryModifica_Click()
  ModificaQuery
End Sub
Private Sub mnuGuidaInfo_Click()
  frmInfo.Show vbModal
End Sub
Private Sub ConnessioneDB()
  On Error GoTo Errors
  sbrStato.SimpleText = "Selezionare il database da aprire"
  sbrStato.Refresh
  With dlgConnessione
    .Filter = "Microsoft Access databases (*.mdb)|*.mdb|Financial Data database|financial data.mdb|"
    .DefaultExt = "fin*.mdb"
    .DialogTitle = "Selezionare il database"
    .Flags = cdlOFNPathMustExist + cdlOFNFileMustExist + cdlOFNExplorer + cdlOFNHideReadOnly
    .FileName = ""
  End With
  Dim strNomeDatabase As String
  dlgConnessione.ShowOpen
  MousePointer = vbHourglass
  strNomeDatabase = dlgConnessione.FileName
  If strNomeDatabase <> "" Then
    sbrStato.SimpleText = "Apertura della connessione al database in corso . . ."
    sbrStato.Refresh
    Set wspFinData = DBEngine.Workspaces(0)
    Set dbFinData = wspFinData.OpenDatabase(strNomeDatabase)
    IntervalloTemporale
    mnuQueryNuova.Enabled = True
    sbrStato.SimpleText = "Connessione al database aperta"
    sbrStato.Refresh
    tlbBarraStrumenti.Buttons(1).Enabled = False
    tlbBarraStrumenti.Buttons(2).Enabled = True
    tlbBarraStrumenti.Buttons(5).Enabled = True
    mnuDatabaseChiudi.Enabled = True
    mnuDatabaseApri.Enabled = False
    mnuQuery.Enabled = True
    mnuGrafico.Enabled = True
  End If
  MousePointer = 0
Errors:
  Select Case Err.Number
    Case 3078
      o = MsgBox("Il database selezionato non è compatibile con Financial Data.", vbExclamation, "Financial Data")
      MousePointer = 0
      sbrStato.SimpleText = "Financial Data 1.0"
      sbrStato.Refresh
      Exit Sub
      Exit Sub
  End Select
End Sub
Sub IntervalloTemporale()
  Dim recDateMinMax As Recordset
  Dim strSQLdateMinMax As String
  strSQLdateMinMax = "SELECT min(data),max(data) FROM Prezzi"
  Set recDateMinMax = dbFinData.OpenRecordset(strSQLdateMinMax, dbOpenSnapshot, dbReadOnly)
  intMin = Val(Format(recDateMinMax(0), "yyyy"))
  intMax = Val(Format(recDateMinMax(1), "yyyy"))
End Sub
Private Sub tlbBarraStrumenti_ButtonClick(ByVal Button As ComctlLib.Button)
  Select Case Button.Key
    Case "ApriConnessione"
      ConnessioneDB
    Case "ImpostaQuery"
      NuovaQuery
    Case "ModificaQuery"
      ModificaQuery
    Case "GraficoIndicatore"
      frmIndicatori.Show
    Case "GraficoSerieStorica"
      GraficoSerie
  End Select
End Sub
Private Sub NuovaQuery()
  ScaricaFormsAttivi
  frmQueryStocks.WindowState = 2
  frmQueryStocks.Show
End Sub
Private Sub ModificaQuery()
  udtFormsLoad.Qstocks = False
  udtFormsLoad.Qdate = False
  udtFormsLoad.Qcampi = False
  udtFormsLoad.Qvalori = False
  ScaricaFormsAttivi
  frmQueryStocks.WindowState = 2
  frmQueryStocks.Show
End Sub
Private Sub GraficoSerie()
  ScaricaFormsAttivi
  frmSerieStoriche.WindowState = 2
  frmSerieStoriche.Show
End Sub
