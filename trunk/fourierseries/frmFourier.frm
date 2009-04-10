VERSION 5.00
Object = "{008BBE7B-C096-11D0-B4E3-00A0C901D681}#1.0#0"; "TEECHART.OCX"
Begin VB.Form frmFourier 
   Caption         =   " Método de Fourier"
   ClientHeight    =   5730
   ClientLeft      =   720
   ClientTop       =   1560
   ClientWidth     =   10635
   Icon            =   "frmFourier.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   101.071
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   187.59
   WindowState     =   2  'Maximized
   Begin TeeChart.TChart tchGrafico 
      Height          =   3735
      Left            =   0
      OleObjectBlob   =   "frmFourier.frx":030A
      TabIndex        =   0
      Top             =   0
      Width           =   4575
   End
   Begin TeeChart.TChart tchEspectro 
      Height          =   3735
      Left            =   4800
      OleObjectBlob   =   "frmFourier.frx":054C
      TabIndex        =   19
      Top             =   0
      Width           =   4335
   End
   Begin VB.CommandButton cmdExibirOcultarHarmonicas 
      Caption         =   "Ocultar Harmônicas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   21
      Top             =   6120
      Width           =   2175
   End
   Begin VB.CommandButton cmdEditarEspectro 
      Caption         =   "Editar Espectro de Linha"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   20
      Top             =   5520
      Width           =   2175
   End
   Begin VB.CommandButton cmdTracar 
      Caption         =   "Traçar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   18
      Top             =   4320
      Width           =   2175
   End
   Begin VB.CommandButton cmdEditarHarmonicas 
      Caption         =   "Editar Harmônicas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9240
      TabIndex        =   17
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Frame fraDados 
      Caption         =   " Dados: "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   3240
      TabIndex        =   5
      Top             =   4320
      Width           =   5655
      Begin VB.OptionButton optHarmonica 
         Caption         =   "Harmônica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   1560
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optPonto 
         Caption         =   "Ponto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   15
         Top             =   1560
         Width           =   855
      End
      Begin VB.TextBox txtComprimento 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         TabIndex        =   12
         Text            =   "3"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtPasso 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         TabIndex        =   10
         Text            =   "0.05"
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox txtAmplitude 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Text            =   "1"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox txtHarmonicas 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Text            =   "5"
         Top             =   480
         Width           =   615
      End
      Begin VB.Label lblAtualizar 
         AutoSize        =   -1  'True
         Caption         =   "Atualizar a cada:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   14
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblComprimento 
         AutoSize        =   -1  'True
         Caption         =   "Comprimento (xPI):"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   13
         Top             =   960
         Width           =   1410
      End
      Begin VB.Label lblPasso 
         AutoSize        =   -1  'True
         Caption         =   "Passo do Ponto:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2640
         TabIndex        =   11
         Top             =   480
         Width           =   1170
      End
      Begin VB.Label lblAmplitude 
         AutoSize        =   -1  'True
         Caption         =   "Amplitude:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   8
         Top             =   960
         Width           =   765
      End
      Begin VB.Label lblHarmonicas 
         AutoSize        =   -1  'True
         Caption         =   "Harmônicas:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   6
         Top             =   480
         Width           =   885
      End
   End
   Begin VB.Frame fraOnda 
      Caption         =   " Tipo de Onda:  "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   4320
      Width           =   2775
      Begin VB.OptionButton optTriangular 
         Caption         =   "Triangular"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton optDenteSerra 
         Caption         =   "Dente de Serra"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton optQuadrada 
         Caption         =   "Quadrada"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   2
         Top             =   600
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "&Ajuda"
      Begin VB.Menu mnuSobre 
         Caption         =   "Sobre ..."
      End
   End
End
Attribute VB_Name = "frmFourier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const PI = 3.14159265358979
Public n As Long, ponto As Long, X As Single, Y As Single, AmplitudeEspectro As Single

Sub EspectroTriangular(Harmonicas As Long, Amplitude As Integer)

tchEspectro.AddSeries scBar
tchEspectro.Series(0).Title = "Espectro de Linha"
tchEspectro.Series(0).Marks.Visible = False
tchEspectro.Series(0).Marks.Style = smsValue
For n = 1 To Harmonicas
    AmplitudeEspectro = ((6 * Amplitude) / (n ^ 2 * PI ^ 2)) * (Sin(n * PI / 2)) * Sin(n * X)
    tchEspectro.Series(0).Add AmplitudeEspectro, CStr(n), tchGrafico.Series(n).SeriesColor
Next

End Sub

Sub Triangular(Harmonicas As Long, Amplitude As Integer, Passo As Single, Comprimento As Single)

'Traçando a onda
tchGrafico.AddSeries scFastLine
tchGrafico.Series(0).Title = "Onda"
ponto = 0
For X = -(Comprimento * PI) To (Comprimento * PI) Step Passo
    Y = ((6 * Amplitude) / (PI ^ 2)) * (Sin(PI / 2)) * Sin(X)
    tchGrafico.Series(0).AddXY X, Y, "", vbBlue
Next

'Traçando a fundamental
tchGrafico.AddSeries scFastLine
tchGrafico.Series(1).Title = "Fundamental"
For X = -(Comprimento * PI) To (Comprimento * PI) Step Passo
    Y = ((6 * Amplitude) / (PI ^ 2)) * (Sin(PI / 2)) * Sin(X)
    tchGrafico.Series(1).AddXY X, Y, "", vbBlue
Next

'Traçando as outras harmônicas
For n = 2 To Harmonicas
tchGrafico.AddSeries scFastLine
tchGrafico.Series(n).Title = n & "ª Harmônica"
ponto = tchGrafico.Series(0).Count
    For X = -(Comprimento * PI) To (Comprimento * PI) Step Passo
        Y = ((6 * Amplitude) / (n ^ 2 * PI ^ 2)) * (Sin(n * PI / 2)) * Sin(n * X)
        tchGrafico.Series(n).AddXY X, Y, "", vbBlue
        tchGrafico.Series(0).YValues.Value(tchGrafico.Series(0).Count - ponto) = tchGrafico.Series(0).YValues.Value(tchGrafico.Series(0).Count - ponto) + Y
        ponto = ponto - 1
        If optPonto.Value Then DoEvents
    Next
If optHarmonica.Value Then DoEvents
Next
EspectroTriangular Harmonicas, Amplitude

End Sub

Sub EspectroDenteSerra(Harmonicas As Long, Amplitude As Integer)

tchEspectro.AddSeries scBar
tchEspectro.Series(0).Title = "Espectro de Linha"
tchEspectro.Series(0).Marks.Visible = False
tchEspectro.Series(0).Marks.Style = smsValue
For n = 1 To Harmonicas
    AmplitudeEspectro = ((-2 * Amplitude) / (n * PI)) * (Cos(n * PI))
    tchEspectro.Series(0).Add AmplitudeEspectro, CStr(n), tchGrafico.Series(n).SeriesColor
Next

End Sub

Sub DenteSerra(Harmonicas As Long, Amplitude As Integer, Passo As Single, Comprimento As Single)

'Traçando a onda
tchGrafico.AddSeries scFastLine
tchGrafico.Series(0).Title = "Onda"
ponto = 0
For X = -(Comprimento * PI) To (Comprimento * PI) Step Passo
    Y = ((-2 * Amplitude) / (PI)) * (Cos(PI)) * Sin(X)
    tchGrafico.Series(0).AddXY X, Y, "", vbBlue
Next

'Traçando a fundamental
tchGrafico.AddSeries scFastLine
tchGrafico.Series(1).Title = "Fundamental"
For X = -(Comprimento * PI) To (Comprimento * PI) Step Passo
    Y = ((-2 * Amplitude) / (PI)) * (Cos(PI)) * Sin(X)
    tchGrafico.Series(1).AddXY X, Y, "", vbBlue
Next

'Traçando as outras harmônicas
For n = 2 To Harmonicas
tchGrafico.AddSeries scFastLine
tchGrafico.Series(n).Title = n & "ª Harmônica"
ponto = tchGrafico.Series(0).Count
    For X = -(Comprimento * PI) To (Comprimento * PI) Step Passo
        Y = ((-2 * Amplitude) / (n * PI)) * (Cos(n * PI)) * Sin(n * X)
        tchGrafico.Series(n).AddXY X, Y, "", vbBlue
        tchGrafico.Series(0).YValues.Value(tchGrafico.Series(0).Count - ponto) = tchGrafico.Series(0).YValues.Value(tchGrafico.Series(0).Count - ponto) + Y
        ponto = ponto - 1
        If optPonto.Value Then DoEvents
    Next
If optHarmonica.Value Then DoEvents
Next
EspectroDenteSerra Harmonicas, Amplitude

End Sub

Sub EspectroQuadrada(Harmonicas As Long, Amplitude As Integer)

tchEspectro.AddSeries scBar
tchEspectro.Series(0).Title = "Espectro de Linha"
tchEspectro.Series(0).Marks.Visible = False
tchEspectro.Series(0).Marks.Style = smsValue
For n = 1 To Harmonicas
    AmplitudeEspectro = ((2 * Amplitude) / (n * PI)) * (1 - Cos(n * PI))
    tchEspectro.Series(0).Add AmplitudeEspectro, CStr(n), tchGrafico.Series(n).SeriesColor
Next

End Sub
Sub Quadrada(Harmonicas As Long, Amplitude As Integer, Passo As Single, Comprimento As Single)

'Traçando a onda
tchGrafico.AddSeries scFastLine
tchGrafico.Series(0).Title = "Onda"
ponto = 0
For X = -(Comprimento * PI) To (Comprimento * PI) Step Passo
    Y = ((2 * Amplitude) / (PI)) * (1 - Cos(PI)) * Sin(X)
    tchGrafico.Series(0).AddXY X, Y, "", vbBlue
Next

'Traçando a fundamental
tchGrafico.AddSeries scFastLine
tchGrafico.Series(1).Title = "Fundamental"
For X = -(Comprimento * PI) To (Comprimento * PI) Step Passo
    Y = ((2 * Amplitude) / (PI)) * (1 - Cos(PI)) * Sin(X)
    tchGrafico.Series(1).AddXY X, Y, "", vbBlue
Next

'Traçando as outras harmônicas
For n = 2 To Harmonicas
tchGrafico.AddSeries scFastLine
tchGrafico.Series(n).Title = n & "ª Harmônica"
ponto = tchGrafico.Series(0).Count
    For X = -(Comprimento * PI) To (Comprimento * PI) Step Passo
        Y = ((2 * Amplitude) / (n * PI)) * (1 - Cos(n * PI)) * Sin(n * X)
        tchGrafico.Series(n).AddXY X, Y, "", vbBlue
        tchGrafico.Series(0).YValues.Value(tchGrafico.Series(0).Count - ponto) = tchGrafico.Series(0).YValues.Value(tchGrafico.Series(0).Count - ponto) + Y
        ponto = ponto - 1
        If optPonto.Value Then DoEvents
    Next
If optHarmonica.Value Then DoEvents
Next
EspectroQuadrada Harmonicas, Amplitude

End Sub

Private Sub cmdEditarEspectro_Click()

tchEspectro.ShowEditor

End Sub

Private Sub cmdEditarHarmonicas_Click()

tchGrafico.ShowEditor

End Sub

Private Sub cmdExibirOcultarHarmonicas_Click()

tchGrafico.Axis.Left.Automatic = False
If cmdExibirOcultarHarmonicas.Caption = "Ocultar Harmônicas" Then
    cmdExibirOcultarHarmonicas.Caption = "Exibir Harmônicas"
Else
    cmdExibirOcultarHarmonicas.Caption = "Ocultar Harmônicas"
End If
For n = 1 To tchGrafico.SeriesCount - 1
    tchGrafico.Series(n).Active = Not tchGrafico.Series(n).Active
Next

End Sub

Private Sub cmdTracar_Click()

cmdExibirOcultarHarmonicas.Enabled = False
cmdEditarHarmonicas.Enabled = False
cmdEditarEspectro.Enabled = False
tchGrafico.RemoveAllSeries
tchGrafico.Axis.Bottom.Automatic = True
tchGrafico.Axis.Left.Automatic = True
tchEspectro.RemoveAllSeries
tchEspectro.Axis.Bottom.Automatic = True
tchEspectro.Axis.Left.Automatic = True
If cmdExibirOcultarHarmonicas.Caption = "Exibir Harmônicas" Then
    cmdExibirOcultarHarmonicas.Caption = "Ocultar Harmônicas"
End If
If optQuadrada.Value Then
    Quadrada CLng(txtHarmonicas.Text), CInt(txtAmplitude.Text), CSng(Val(txtPasso.Text)), CSng(txtComprimento.Text / 2)
ElseIf optDenteSerra.Value Then
    DenteSerra CLng(txtHarmonicas.Text), CInt(txtAmplitude.Text), CSng(Val(txtPasso.Text)), CSng(txtComprimento.Text / 2)
ElseIf optTriangular.Value Then
    Triangular CLng(txtHarmonicas.Text), CInt(txtAmplitude.Text), CSng(Val(txtPasso.Text)), CSng(txtComprimento.Text / 2)
End If
cmdExibirOcultarHarmonicas.Enabled = True
cmdEditarHarmonicas.Enabled = True
cmdEditarEspectro.Enabled = True

End Sub

Private Sub Form_Activate()

tchGrafico.Top = 2
tchGrafico.Width = 2 * (frmFourier.ScaleWidth / 3) - 10
tchGrafico.Left = 2
tchGrafico.Height = (frmFourier.ScaleHeight / 2) + 10

tchEspectro.Top = 2
tchEspectro.Width = frmFourier.ScaleWidth - tchGrafico.Width - 6
tchEspectro.Left = tchGrafico.Width + 4
tchEspectro.Height = tchGrafico.Height

fraOnda.Left = 4
fraOnda.Top = tchGrafico.Height + 10

fraDados.Left = (frmFourier.ScaleWidth / 2) - (fraDados.Width / 2) + 4
fraDados.Top = fraOnda.Top

cmdTracar.Left = frmFourier.ScaleWidth - cmdEditarHarmonicas.Width - 6
cmdEditarHarmonicas.Left = cmdTracar.Left
cmdEditarEspectro.Left = cmdTracar.Left
cmdExibirOcultarHarmonicas.Left = cmdTracar.Left

cmdTracar.Top = fraOnda.Top + 2
cmdEditarHarmonicas.Top = cmdTracar.Top + cmdTracar.Height + 4
cmdEditarEspectro.Top = cmdEditarHarmonicas.Top + cmdEditarHarmonicas.Height + 4
cmdExibirOcultarHarmonicas.Top = cmdEditarEspectro.Top + cmdEditarEspectro.Height + 4

End Sub

Private Sub mnuSobre_Click()

MsgBox "Método de Fourier Versão 1.0" & vbCrLf & vbCrLf & "Desenvolvido por:" & vbCrLf & "Ricardo Iramar dos Santos" & vbCrLf & "E-mail:" & vbCrLf & "riramar@terra.com.br", vbOKOnly + vbInformation, "Sobre"

End Sub
