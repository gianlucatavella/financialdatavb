VERSION 5.00
Begin VB.Form frmCandleStick 
   AutoRedraw      =   -1  'True
   Caption         =   "Grafico della serie storica: Candle Stick"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11625
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   493
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   775
   WindowState     =   2  'Maximized
   Begin VB.Label Label2 
      Caption         =   "Tempo"
      Height          =   255
      Left            =   5760
      TabIndex        =   2
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Valori"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblTitolo 
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
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   7095
   End
End
Attribute VB_Name = "frmCandleStick"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
  MDIFd.mnuGraficoChiudi.Enabled = True
  udtFormsLoad.CandleStick = True
  MDIFd.sbrStato.Visible = False
  Cls
  Line (70, 30)-(70, 432)
  Line (70, 432)-(775, 432)
  Call RiempiCandleStick
End Sub
Private Sub RiempiCandleStick()
  Dim intCol As Integer
  Dim intRig As Integer
  Dim intCandleMezzo As Integer
  Dim i As Integer
  Dim intColore As Integer
  Dim intX As Integer
  Dim intN As Integer
  Dim strAnni As String
  intCol = 72
  intCandleMezzo = Int((intCandleLargh / 2) + 1)
  intX = 60
  intN = 1
  For i = 1 To UBound(udtCandle) - 1
    If udtCandle(i).Apertura > udtCandle(i).Ultimo Then
      intColore = 15
    Else
      intColore = 0
    End If
    Line ((intCol + intCandleMezzo), udtCandle(i).Alto)-((intCol + intCandleMezzo), udtCandle(i).Basso), QBColor(0)
    Line (intCol, udtCandle(i).Apertura)-((intCol + intCandleLargh), udtCandle(i).Ultimo), QBColor(intColore), BF
    Line ((intCol + intCandleMezzo), 432)-((intCol + intCandleMezzo), 436)
    If intCol >= (intX * intN) Then
      intN = intN + 1
      CurrentX = intCol + intCandleMezzo
      CurrentY = 450
      ForeColor = QBColor(1)
      Print Format(udtCandle(i).Data, "dd/mm")
      Line ((intCol + intCandleMezzo), 432)-((intCol + intCandleMezzo), 445)
      ForeColor = QBColor(0)
    End If
    intCol = intCol + (intCandleLargh + intCandleDist)
  Next
  For intRig = 32 To 412 Step 20
    ForeColor = QBColor(1)
    Line (70, intRig)-(65, intRig)
    CurrentX = 24
    CurrentY = intRig - 7
    Print Format(dblCandleMin + (430 - intRig) * dblCandleUnit, "0.000")
  Next
  If frmSerieStoriche.cmbY1 <> frmSerieStoriche.cmbY2 Then
    strAnni = frmSerieStoriche.cmbY1 & "-" & LTrim(frmSerieStoriche.cmbY2)
  Else
    strAnni = frmSerieStoriche.cmbY1
  End If
  lblTitolo.Caption = frmSerieStoriche.cmbStocks & ": " & strAnni
End Sub
