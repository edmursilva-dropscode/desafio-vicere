VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{DC81D4AD-48D8-4DD6-A8B5-228CB11C1826}#1.0#0"; "prjXTab.ocx"
Begin VB.Form frmCorretor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Corretor"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6645
   LinkTopic       =   "frmCorretor"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2370
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   360
      Left            =   5520
      TabIndex        =   9
      Top             =   1875
      Width           =   1005
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   360
      Left            =   4455
      TabIndex        =   8
      Top             =   1905
      Width           =   1005
   End
   Begin prjXTab.XTab xtbCliente 
      Height          =   960
      Left            =   135
      TabIndex        =   1
      Top             =   810
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1693
      TabCount        =   1
      TabCaption(0)   =   "  Corretor  "
      TabContCtrlCnt(0)=   2
      Tab(0)ContCtrlCap(1)=   "fraTab1"
      Tab(0)ContCtrlCap(2)=   "lblCodigo"
      TabStyle        =   1
      TabTheme        =   1
      ShowFocusRect   =   0   'False
      ActiveTabBackStartColor=   16514555
      InActiveTabBackStartColor=   16777215
      InActiveTabBackEndColor=   15397104
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   10198161
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      Begin VB.TextBox Text1 
         Height          =   1575
         Left            =   -74940
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   5080
      End
      Begin VB.Frame fraTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   420
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   435
         Width           =   6180
         Begin VB.TextBox txtNome 
            Height          =   315
            Left            =   1185
            MaxLength       =   40
            TabIndex        =   3
            Tag             =   "0"
            Top             =   0
            Width           =   4935
         End
         Begin VB.Label lblNome 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nome:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   0
            Left            =   15
            TabIndex        =   4
            Top             =   75
            Width           =   570
         End
      End
      Begin VB.Label lblCodigo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1020
         TabIndex        =   6
         Top             =   15
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lvwCorretor 
      Height          =   360
      Left            =   2790
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   635
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IdCliente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   12
      Left            =   75
      Picture         =   "frmCorretor.frx":0000
      Top             =   75
      Width           =   480
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17955
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Left            =   -1590
      Picture         =   "frmCorretor.frx":08CA
      Top             =   660
      Width           =   10740
   End
End
Attribute VB_Name = "frmCorretor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Dim vop_CorretoresNegocios As New clsCorretoresNegocios

'Variaveis de controle do form
Private vbp_Corretor As Boolean                             'Verifica uma inclusao ou alteracao


'Eventos
Private Sub Form_Activate()
   
   Me.Refresh
   
End Sub

Public Sub Form_Load()
    
    vbp_Corretor = False              'Verifica uma inclusao ou alteracao
    
    'Inicializa entrada e saida
    Call InicializaCorretor
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo TrataErros

    'Tecla de sair do form
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
TrataErros:
    If Err.Number <> 0 Then
        Err.Clear
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        KeyAscii = 0
        Unload Me
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   Set frmCorretor = Nothing
   
End Sub

Private Sub txtNome_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
End Sub


Private Sub cmdGravar_Click()
Dim vsp_Mensagem As String

   'Valida mensagem
   If vbp_Corretor = False Then
      vsp_Mensagem = "Confirma a Inclusão ?"
   Else
      vsp_Mensagem = "Confirma a Alteração ?"
   End If

   'Valida entrada de dados
   If VerCampos = False Then Exit Sub
   
   If MsgBox(vsp_Mensagem, vbQuestion + vbYesNo, "Confirme !") = vbYes Then
   
      Set vop_CorretoresNegocios = New clsCorretoresNegocios
          vop_CorretoresNegocios.IdCorretor = lblCodigo.Caption
          vop_CorretoresNegocios.Nome = txtNome.text
          If vbp_Corretor = False Then
             If vop_CorretoresNegocios.IncluirCorretor(lvwCorretor) = True Then
                MsgBox "Corretor cadastrado com sucesso !", vbExclamation, "Corretor"
             End If
          Else
             If vop_CorretoresNegocios.AlterarCorretor(lvwCorretor) = True Then
                MsgBox "Corretor alterado com sucesso !", vbExclamation, "Corretor"
             End If
          End If
      Set vop_CorretoresNegocios = Nothing
      
   End If
   
   'Atualiza grid
   Call frmListaCorretores.CarregarGrid
   
   'Inicializa entrada e saida
   Call InicializaCorretor
   
   'Valida entrada de dados
   If vbp_Corretor = True Then
      Call cmdFechar_Click
   End If
   
End Sub

Private Sub cmdFechar_Click()
    Unload Me
End Sub



'Funcoes
Function Editar(ByVal pIdCorretor As Integer) As Boolean
    
    'Verifica uma inclusao ou alteracao do Corretor
    vbp_Corretor = True
    'Controle de exibicao
    lblCodigo.Visible = True
    
    Set vop_CorretoresNegocios = New clsCorretoresNegocios
    
        If vop_CorretoresNegocios.PesquisarCorretor(lvwCorretor, pIdCorretor, 0) = True Then
           lblCodigo.Caption = pIdCorretor
           txtNome.text = vop_CorretoresNegocios.Nome
        Else
            MsgBox "Não foi possível encontrar o Corretor !", vbCritical, "Corretor"
        End If
          
    Set vop_CorretoresNegocios = Nothing
    
    Me.Show vbModal

End Function

Function VerCampos() As Boolean
    
    If Trim$(txtNome.text) = Empty Then
        MsgBox "Informe o nome do Corretor !", vbExclamation, "Corretor"
        'txtNome.SetFocus
        VerCampos = False
        Exit Function
    End If
    
    VerCampos = True

End Function

Private Function InicializaCorretor() As Boolean

    'Limpa entrada de dados
    Call LimpaCampos(Me)
    
End Function

Function RemoverPontosEUnderscores(ByVal texto As String) As String
    Dim resultado As String
    Dim i As Integer

    ' Inicializa a string de resultado
    resultado = ""

    ' Itera por cada caractere na string de entrada
    For i = 1 To Len(texto)
        Dim caractereAtual As String
        caractereAtual = Mid(texto, i, 1)

        ' Verifica se o caractere não é '.' ou '_'
        If caractereAtual <> "." And caractereAtual <> "_" Then
            resultado = resultado & caractereAtual ' Adiciona o caractere à string de resultado
        End If
    Next i

    ' Retorna a string sem os caracteres indesejad
    RemoverPontosEUnderscores = resultado
End Function
