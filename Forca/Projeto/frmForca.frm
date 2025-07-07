VERSION 5.00
Begin VB.Form frmForca 
   AutoRedraw      =   -1  'True
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forca_2506"
   ClientHeight    =   12225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14550
   Icon            =   "frmForca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12225
   ScaleWidth      =   14550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraForca 
      BackColor       =   &H00404040&
      Enabled         =   0   'False
      Height          =   11895
      Left            =   360
      TabIndex        =   4
      Top             =   120
      Width           =   9615
      Begin VB.TextBox txtLetra 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   480
         Left            =   600
         MaxLength       =   1
         TabIndex        =   7
         Text            =   "A"
         Top             =   840
         Width           =   615
      End
      Begin VB.PictureBox picForca 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00808080&
         DrawWidth       =   5
         ForeColor       =   &H00FF0000&
         Height          =   5475
         Left            =   1560
         ScaleHeight     =   5445
         ScaleWidth      =   7050
         TabIndex        =   6
         Top             =   2520
         Width           =   7080
      End
      Begin VB.ListBox lstLetras 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   10500
         ItemData        =   "frmForca.frx":0442
         Left            =   600
         List            =   "frmForca.frx":0444
         TabIndex        =   5
         Top             =   1320
         Width           =   615
      End
      Begin VB.Image imgAcerto 
         Height          =   480
         Left            =   8880
         Picture         =   "frmForca.frx":0446
         Stretch         =   -1  'True
         Top             =   1080
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblCategoria 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Categorias: Fruta - Carro - Roupa - Flor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2520
         TabIndex        =   24
         Top             =   240
         Width           =   6135
      End
      Begin VB.Label lblLetra 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Digite uma letra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label lblErros 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1560
         TabIndex        =   22
         Top             =   8040
         Width           =   2295
      End
      Begin VB.Label lblletrasErradas 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1560
         TabIndex        =   21
         Top             =   1920
         Width           =   7095
      End
      Begin VB.Label lblLetraSel 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   0
         Left            =   1560
         TabIndex        =   20
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblLetraSel 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   1
         Left            =   2280
         TabIndex        =   19
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblLetraSel 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   2
         Left            =   3000
         TabIndex        =   18
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblLetraSel 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   3
         Left            =   3720
         TabIndex        =   17
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblLetraSel 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   4
         Left            =   4440
         TabIndex        =   16
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblLetraSel 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   5
         Left            =   5160
         TabIndex        =   15
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblLetraSel 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   6
         Left            =   5880
         TabIndex        =   14
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblLetraSel 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   7
         Left            =   6600
         TabIndex        =   13
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblLetraSel 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   8
         Left            =   7320
         TabIndex        =   12
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblLetraSel 
         Alignment       =   2  'Center
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   9
         Left            =   8040
         TabIndex        =   11
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblRodada 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1560
         TabIndex        =   10
         Top             =   8760
         Width           =   2295
      End
      Begin VB.Label lblJogador2 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Jogador 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1560
         TabIndex        =   9
         Top             =   10200
         Width           =   2295
      End
      Begin VB.Label lblJogador1 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Jogador 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1560
         TabIndex        =   8
         Top             =   9480
         Width           =   2295
      End
   End
   Begin VB.Timer tmrTempo 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   11400
      Top             =   720
   End
   Begin VB.ListBox lstPalavras 
      Height          =   4935
      Left            =   10080
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton cmdReiniciar 
      Caption         =   "&Reiniciar Jogo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      TabIndex        =   2
      Top             =   8880
      Width           =   1575
   End
   Begin VB.CommandButton cmdDicionario 
      Caption         =   "&Dicionario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      TabIndex        =   1
      Top             =   9840
      Width           =   1575
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "&Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      TabIndex        =   0
      Top             =   10800
      Width           =   1575
   End
End
Attribute VB_Name = "frmForca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim palavra As String
Dim erros As Integer
Dim tempoForca As Integer
Dim rodadas As Integer
Dim caminhoBD As String
Dim Jogador1, Jogador2 As Integer

Private Sub cmdDicionario_Click()
  Dim aux As String
  aux = UCase(InputBox("Digte uma palavra (máximo 10 letras):", "Dicionário"))
  If Len(aux) <= 10 Then
    palavra = aux
    inicializarJogo
  Else
    MsgBox "Favor digitar uma palavra com máximo 10 letras"
  End If
  
End Sub

Private Sub cmdSair_Click()
    End
End Sub

Private Sub Form_Activate()
    If Trim(palavra) = "" Then
        ListarPalavras
        SorteiaPalavra
        inicializarJogo
        tmrTempo.Enabled = True
    End If
    ListaLetras
End Sub

Private Sub Form_Load()
    'palavra = "SAPATO"
    caminhoBD = "C:\CURSOS_Programador - 2025\Visual Basic 5 - Curso\forca2025\dicionario.txt"
    'caminhoBD = CurDir & "\dicionario.txt"
   
    erros = 0
    txtLetra.Text = ""
    picForca.Cls
    limparLetras
    lblErros.Caption = "Erros: 0"
End Sub

Private Sub limparLetras()
    Dim i As Integer
    For i = 0 To lblLetraSel().Count - 1
        lblLetraSel(i).Caption = "?" ' ou "_" se preferir mostrar underline
    Next i
End Sub

Private Sub lstLetras_Click()
If Trim(lstLetras) <> "" Then ' Tecla Enter
        txtLetra.Text = Trim(lstLetras)
        verificarLetra
        txtLetra.Text = ""
    End If
End Sub

Private Sub tmrTempo_Timer()
    'Dim i As Integer
    tempoForca = tempoForca + 1
    'For i = 1 To 10
     If tempoForca <= 10 Then
        desenharForca tempoForca
     Else
        tempoForca = 0
        tmrTempo.Enabled = False
        picForca.Cls
        fraForca.Enabled = True
        txtLetra.SetFocus
     End If
    'Next i
End Sub

Private Sub txtLetra_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

    If KeyAscii = 13 And Trim(palavra) <> "" Then ' Tecla Enter
        verificarLetra
        txtLetra.Text = ""
    End If
End Sub

Private Sub verificarLetra()
    Dim achou As Boolean
    Dim i As Integer
    Dim aux As String
    achou = False
    imgAcerto.Visible = False
    If Trim(palavra) = "" Then
        MsgBox "erro palavra nula"
        Exit Sub
    End If
    
    For i = 1 To Len(palavra)
        aux = Mid(palavra, i, 1)
        
        If aux = txtLetra.Text Then
            lblLetraSel(i - 1).Caption = txtLetra.Text
            'Beep
            imgAcerto.Visible = True
            achou = True
        End If
    Next i

    If Not achou Then
        erros = erros + 1
        lblErros.Caption = "Erros: " & erros
        lblletrasErradas.Caption = lblletrasErradas.Caption & txtLetra.Text & ", "
        desenharForca erros
    End If

    If erros >= 10 Then
        MsgBox "Fim de jogo! A palavra era: " & palavra, vbExclamation, "Você perdeu"
        Jogador2 = Jogador2 + 1
        lblJogador2.Caption = "Jogador 2: " & Jogador2 & " pts"
        SorteiaPalavra
        inicializarJogo
        Exit Sub
    End If

    If palavraCompleta() Then
        MsgBox "Parabéns! Você acertou a palavra!", vbInformation, "Vitória"
         Jogador1 = Jogador1 + 1
        lblJogador1.Caption = "Jogador 1: " & Jogador1 & " pts"
        SorteiaPalavra
        inicializarJogo
    End If
End Sub

Private Function palavraCompleta() As Boolean
    Dim i As Integer
    For i = 0 To Len(palavra) - 1
        If lblLetraSel(i).Caption = "?" Then
            palavraCompleta = False
            Exit Function
        End If
    Next i
    palavraCompleta = True
End Function

Private Sub inicializarJogo()
    erros = 0
    lblErros.Caption = "Erros: 0"
    lblletrasErradas.Caption = ""
    picForca.Cls
    limparLetras
    rodadas = rodadas + 1
    lblRodada.Caption = "Rodadas: " & rodadas
End Sub

Private Sub cmdReiniciar_Click()
    fraForca.Enabled = False
    SorteiaPalavra
    inicializarJogo
    picForca.Cls
    tmrTempo.Enabled = True
End Sub

Private Sub desenharForca(etapa As Integer)
    Select Case etapa
        Case 1 ' Base
            picForca.Line (10 * 19, 200 * 19)-(100 * 19, 200 * 19)
        Case 2 ' Poste vertical
            picForca.Line (55 * 19, 200 * 19)-(55 * 19, 50 * 19)
        Case 3 ' Haste superior
            picForca.Line (55 * 19, 50 * 19)-(150 * 19, 50 * 19)
        Case 4 ' Corda
            picForca.Line (150 * 19, 50 * 19)-(150 * 19, 70 * 19)
        Case 5 ' Cabeça
            picForca.Circle (150 * 19, 90 * 19), 20 * 19
        Case 6 ' Corpo
            picForca.Line (150 * 19, 110 * 19)-(150 * 19, 160 * 19)
        Case 7 ' Braço esquerdo
            picForca.Line (150 * 19, 120 * 19)-(130 * 19, 140 * 19)
        Case 8 ' Braço direito
            picForca.Line (150 * 19, 120 * 19)-(170 * 19, 140 * 19)
        Case 9 ' Perna esquerda
            picForca.Line (150 * 19, 160 * 19)-(130 * 19, 190 * 19)
        Case 10 ' Perna direita
            picForca.Line (150 * 19, 160 * 19)-(170 * 19, 190 * 19)
    End Select
End Sub

Public Sub SorteiaPalavra()
    Dim tamanho, i As Integer
    Dim aux, trocaLetra As String
    palavra = ""
    Randomize
    tamanho = lstPalavras.ListCount
    i = Int(tamanho * Rnd)
    trocaLetra = lstPalavras.List(i)
    
    For i = 1 To Len(trocaLetra)
        aux = Mid(trocaLetra, i, 1)
        If aux = "Ã" Or aux = "Á" Then aux = "A"
        If aux = "Ê" Or aux = "É" Then aux = "E"
        If aux = "Í" Then aux = "I"
        If aux = "Õ" Or aux = "Ô" Or aux = "Ó" Then aux = "O"
        If aux = "Ú" Then aux = "U"
        palavra = palavra & aux
    Next i
    If Len(palavra) > 10 Then palavra = Left(palavra, 10)

    'MsgBox Len(palavra) & " - " & palavra
End Sub

Public Sub ListarPalavras()
    Dim i, tamanho As Integer
    Dim aux As String
    
    lstPalavras.Clear

    Open caminhoBD For Input As #1
    Do While Not EOF(1)
        Input #1, aux
        lstPalavras.AddItem UCase(aux)
    Loop
    Close #1
    
End Sub

Public Sub ListaLetras()
    Dim i As Integer
    For i = 65 To 90
        lstLetras.AddItem " " & Chr(i)
    Next i
    
End Sub
