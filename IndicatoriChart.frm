VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmIndicatoriChart 
   Caption         =   "Indicatori descrittivi tipici del mercato azionario: grafici"
   ClientHeight    =   7335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   11535
   Begin MSChart20Lib.MSChart mscIndicatori 
      Height          =   6495
      Left            =   0
      OleObjectBlob   =   "IndicatoriChart.frx":0000
      TabIndex        =   5
      Top             =   120
      Width           =   11535
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
      Height          =   495
      Left            =   7320
      TabIndex        =   2
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSuccessivo 
      Caption         =   "  &Successivo           >>"
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
      Left            =   10440
      TabIndex        =   0
      Top             =   6720
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdPrecedente 
      Caption         =   " &Precedente        <<"
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
      Left            =   9120
      TabIndex        =   1
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblGraficoNonDisp 
      Caption         =   "        Grafico non disponibile"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3480
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label lblTitoloChart 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   6960
      Width           =   6615
   End
End
Attribute VB_Name = "frmIndicatoriChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intNumeroStocks As Integer
Dim blnAmpiezza As Boolean
Private Sub cmdPrecedente_Click()
  intIndexS = intIndexS - 1
  cmdSuccessivo.Visible = True
  If intIndexS = 1 Then
    cmdPrecedente.Visible = False
  End If
  blnAmpiezza = True
  Call RiempiGrafico(udtDynChart(intIndexS).Nome, udtDynChart(intIndexS).Numero)
  If blnAmpiezza = False Then
    mscIndicatori.Visible = False
    lblGraficoNonDisp.Caption = "Grafico di " & udtDynChart(intIndexS).Nome & " non disponibile"
    lblGraficoNonDisp.Visible = True
  Else
    mscIndicatori.Visible = True
    lblGraficoNonDisp.Visible = False
  End If
End Sub
Private Sub cmdChiudi_Click()
  MDIFd.sbrStato.Visible = True
  MDIFd.mnuGraficoChiudi.Enabled = False
  frmIndicatoriChart.Hide
  Unload frmIndicatoriChart
End Sub
Private Sub cmdSuccessivo_Click()
  intIndexS = intIndexS + 1
  cmdPrecedente.Visible = True
  If intIndexS = intNumeroStocks Then
    cmdSuccessivo.Visible = False
  End If
  blnAmpiezza = True
  Call RiempiGrafico(udtDynChart(intIndexS).Nome, udtDynChart(intIndexS).Numero)
  If blnAmpiezza = False Then
    mscIndicatori.Visible = False
    lblGraficoNonDisp.Caption = "Grafico di " & udtDynChart(intIndexS).Nome & " non disponibile"
    lblGraficoNonDisp.Visible = True
  Else
    mscIndicatori.Visible = True
    lblGraficoNonDisp.Visible = False
  End If
End Sub
Private Sub Form_Activate()
  udtFormsLoad.Ichart = True
  MDIFd.mnuGraficoChiudi.Enabled = True
End Sub
Private Sub Form_Load()
  udtFormsLoad.Ichart = True
  MDIFd.mnuGraficoChiudi.Enabled = True
  MDIFd.sbrStato.Visible = False
  intNumeroStocks = UBound(udtDynChart)
  If intNumeroStocks > 1 Then
    cmdSuccessivo.Visible = True
  End If
  blnAmpiezza = True
  Call RiempiGrafico(udtDynChart(intIndexS).Nome, udtDynChart(intIndexS).Numero)
  If blnAmpiezza = False Then
    mscIndicatori.Visible = False
    lblGraficoNonDisp.Caption = "Grafico di " & udtDynChart(intIndexS).Nome & " non disponibile"
    lblGraficoNonDisp.Visible = True
  End If
End Sub
Private Sub RiempiGrafico(strStock As String, intStock As Long)
  On Error GoTo Subscript
  Dim vntDynMSChart() As Variant
  Dim i As Long
  Dim j As Long
  Dim intAmpiezzaChart As Integer
  Dim lngInizio As Long
  If intIndexS = 1 Then
    lngInizio = 1
  Else
    lngInizio = udtDynChart(intIndexS - 1).Numero + 1
  End If
  If frmIndicatori.cmbIndicatori.Text <> "Indice Stocastico" Then
    For i = lngInizio To udtDynChart(intIndexS).Numero
      If dblDynIndicatore(i - 1) <> -1 Then
        intAmpiezzaChart = intAmpiezzaChart + 1
      End If
    Next
    For i = 1 To UBound(udtDynCampi)
      If udtDynCampi(i).Nome = strCampo Then
        lblTitoloChart = "Campo descritto: " & udtDynCampi(i).Descrizione
        Exit For
      End If
    Next
    mscIndicatori.TitleText = frmIndicatori.cmbIndicatori.Text & ":  " & udtDynChart(intIndexS).Nome
    If intAmpiezzaChart <= 2 Then
      blnAmpiezza = False
      Exit Sub
    End If
    ReDim vntDynMSChart(1 To intAmpiezzaChart, 1 To 2)
    recQuery.MoveFirst
    recQuery.Move lngInizio
    i = lngInizio
    j = 1
    Do Until i = udtDynChart(intIndexS).Numero
      If dblDynIndicatore(i - 1) <> -1 Then
        j = j + 1
        vntDynMSChart(j, 1) = Format(recQuery(1), "dd/mm/yy") & "        "
        vntDynMSChart(j, 2) = dblDynIndicatore(i - 1)
      End If
      i = i + 1
      recQuery.MoveNext
    Loop
    Select Case frmIndicatori.cmbIndicatori.Text
      Case "Rate of Change"
        vntDynMSChart(1, 2) = "ROC"
      Case "Media Mobile"
        vntDynMSChart(1, 2) = "MA"
      Case "Relative Strenght Index"
        vntDynMSChart(1, 2) = "RSI"
      Case Else
        vntDynMSChart(1, 2) = " "
     End Select
  Else
    For i = lngInizio To udtDynChart(intIndexS).Numero
      If dblDynStocastico(i - 1).K <> -1 Then
        intAmpiezzaChart = intAmpiezzaChart + 1
      End If
    Next
    mscIndicatori.TitleText = frmIndicatori.cmbIndicatori.Text & ":  " & udtDynChart(intIndexS).Nome
    If intAmpiezzaChart <= 2 Then
       blnAmpiezza = False
      Exit Sub
    End If
    lblTitoloChart = ""
    ReDim vntDynMSChart(1 To intAmpiezzaChart, 1 To 3)
    recQuery.MoveFirst
    i = lngInizio
    j = 1
    Do Until i = udtDynChart(intIndexS).Numero
      If dblDynStocastico(i - 1).K <> -1 Then
        j = j + 1
        vntDynMSChart(j, 1) = Format(recQuery(1), "dd/mm/yy") & "        "
        vntDynMSChart(j, 2) = dblDynStocastico(i - 1).K
        If i > 3 Then
          vntDynMSChart(j, 3) = dblDynStocastico(i - 1).D
        End If
      End If
      i = i + 1
      recQuery.MoveNext
    Loop
    vntDynMSChart(1, 2) = "%K"
    vntDynMSChart(1, 3) = "%D"
  End If
  mscIndicatori = vntDynMSChart
Subscript:
  Select Case Err.Number
    Case 9
      Resume Next
  End Select
End Sub
