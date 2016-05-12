Attribute VB_Name = "Module"
Public strSQL As String
Public strSQLstocks As String
Public strSQLdata As String
Public strSQLvalori As String
Public strDynCampi() As String
Public recQuery As Recordset
Public dbFinData As Database
Public wspFinData As Workspace
Public dblDynIndicatore() As Double
Private Type StocksChart
  Nome As String
  Numero As Long
  Ampiezza As Integer
End Type
Public udtDynChart() As StocksChart
Public intMin As Integer
Public intMax As Integer
Private Type Campi
  Nome As String
  Field As Integer
  Descrizione As String
End Type
Public udtDynCampi() As Campi
Public strCampo As String
Public intAmpGrafico As Integer
Private Type Stocastico
  CL5 As Double
  H5L5 As Double
  K As Double
  D As Double
End Type
Public dblDynStocastico() As Stocastico
Public intIndexS As Integer
Private Type FormsCaricati
  Qstocks As Boolean
  Qcampi As Boolean
  Qdate As Boolean
  Qvalori As Boolean
  Qrisultati As Boolean
  Icalcola As Boolean
  Irisultati As Boolean
  Ichart As Boolean
  Sstoriche As Boolean
  SstoricheChart As Boolean
  CandleStick As Boolean
End Type
Public udtFormsLoad As FormsCaricati
Public blnConferma As Boolean
Private Type CandleStick
  Data As Date
  Apertura As Integer
  Alto As Integer
  Basso As Integer
  Ultimo As Integer
End Type
Public udtCandle() As CandleStick
Public intCandleLargh As Integer
Public intCandleDist As Integer
Public dblCandleUnit As Double
Public dblCandleMin As Double
Public Sub ScaricaFormsAttivi()
    If udtFormsLoad.Sstoriche = True Then
      frmSerieStoriche.Hide
      Unload frmSerieStoriche
      udtFormsLoad.Sstoriche = False
    End If
    If udtFormsLoad.SstoricheChart = True Then
      frmSerieStoricheChart.Hide
      Unload frmSerieStoricheChart
      udtFormsLoad.SstoricheChart = False
    End If
    If udtFormsLoad.CandleStick = True Then
      frmCandleStick.Hide
      Unload frmCandleStick
      udtFormsLoad.CandleStick = False
    End If
    If udtFormsLoad.Qstocks = True Then
      frmQueryStocks.Hide
      Unload frmQueryStocks
      udtFormsLoad.Qstocks = False
    End If
    If udtFormsLoad.Qcampi = True Then
      frmQueryCampi.Hide
      Unload frmQueryCampi
      udtFormsLoad.Qcampi = False
    End If
    If udtFormsLoad.Qdate = True Then
      frmQueryDate.Hide
      Unload frmQueryDate
      udtFormsLoad.Qdate = False
    End If
    If udtFormsLoad.Qrisultati = True Then
      Unload frmQueryRisultati
      udtFormsLoad.Qrisultati = False
    End If
    If udtFormsLoad.Irisultati = True Then
      Unload frmIndicatoriRis
      udtFormsLoad.Irisultati = False
    End If
    If udtFormsLoad.Ichart = True Then
      frmIndicatoriChart.Hide
      Unload frmIndicatoriChart
      udtFormsLoad.Ichart = False
    End If
    If udtFormsLoad.Icalcola = True Then
      frmIndicatori.Hide
      Unload frmIndicatori
      udtFormsLoad.Icalcola = False
    End If
    If udtFormsLoad.Qvalori = True Then
      frmQueryValori.Hide
      Unload frmQueryValori
      udtFormsLoad.Qvalori = False
    End If
    MDIFd.mnuQueryModifica.Enabled = False
    MDIFd.mnuGraficoIndicatori.Enabled = False
    MDIFd.tlbBarraStrumenti.Buttons(3).Enabled = False
    MDIFd.tlbBarraStrumenti.Buttons(4).Enabled = False
    MDIFd.mnuGraficoChiudi.Enabled = False
    MDIFd.sbrStato.Visible = False
End Sub
