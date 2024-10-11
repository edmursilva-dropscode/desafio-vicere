VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Object = "{DC81D4AD-48D8-4DD6-A8B5-228CB11C1826}#1.0#0"; "prjXTab.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTransacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Transação"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9465
   LinkTopic       =   "frmTransacao"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   9465
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin MSComctlLib.ListView lvwCorretor 
      Height          =   360
      Left            =   2880
      TabIndex        =   27
      Top             =   3465
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
         Text            =   "IdCorretor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   360
      Left            =   75
      TabIndex        =   19
      Top             =   3450
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   360
      Left            =   8355
      TabIndex        =   21
      Top             =   3465
      Width           =   1005
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   360
      Left            =   7290
      TabIndex        =   20
      Top             =   3465
      Width           =   1005
   End
   Begin MSACAL.Calendar calData 
      Height          =   315
      Left            =   3255
      TabIndex        =   4
      Top             =   1185
      Width           =   2280
      _Version        =   524288
      _ExtentX        =   4022
      _ExtentY        =   556
      _StockProps     =   1
      BackColor       =   -2147483633
      Year            =   2024
      Month           =   9
      Day             =   18
      DayLength       =   1
      MonthLength     =   2
      DayFontColor    =   0
      FirstDay        =   1
      GridCellEffect  =   2
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.26
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12.01
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin prjXTab.XTab xtbCliente 
      Height          =   2565
      Left            =   45
      TabIndex        =   0
      Top             =   795
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   4524
      TabCount        =   1
      TabCaption(0)   =   "  Transação  "
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
         TabIndex        =   25
         Top             =   360
         Width           =   5080
      End
      Begin VB.Frame fraTab 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         ForeColor       =   &H80000008&
         Height          =   2010
         Index           =   1
         Left            =   180
         TabIndex        =   24
         Top             =   420
         Width           =   8970
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   5025
            TabIndex        =   29
            Text            =   "Text3"
            Top             =   1695
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   5010
            TabIndex        =   28
            Text            =   "Text2"
            Top             =   1350
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.CheckBox chkAtivo 
            Caption         =   "Check1"
            Height          =   225
            Left            =   8520
            TabIndex        =   16
            Top             =   1245
            Width           =   270
         End
         Begin VB.ComboBox cmbLocalizarCidade 
            Height          =   315
            ItemData        =   "frmTransacao.frx":0000
            Left            =   1110
            List            =   "frmTransacao.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   1635
            Width           =   2505
         End
         Begin VB.ComboBox cmbLocalizarEstado 
            Height          =   315
            ItemData        =   "frmTransacao.frx":001C
            Left            =   1110
            List            =   "frmTransacao.frx":0026
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1230
            Width           =   2040
         End
         Begin VB.CommandButton cmdPesquisarCorretor 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3225
            TabIndex        =   11
            Top             =   825
            Width           =   345
         End
         Begin VB.TextBox txtNumeroCorretor 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1110
            MaxLength       =   16
            TabIndex        =   10
            Tag             =   "0"
            Text            =   "0"
            Top             =   825
            Width           =   2055
         End
         Begin VB.CommandButton cmdPesquisar 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   3
            Left            =   3225
            TabIndex        =   7
            Top             =   390
            Width           =   345
         End
         Begin MSMask.MaskEdBox msbCPF 
            Height          =   315
            Left            =   1125
            TabIndex        =   6
            Top             =   420
            Width           =   2025
            _ExtentX        =   3572
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   14
            Mask            =   "###.###.###-##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblAtivo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ativo:"
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
            Index           =   1
            Left            =   7845
            TabIndex        =   15
            Top             =   1260
            Width           =   510
         End
         Begin VB.Label lblNomeCorretor 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   3660
            TabIndex        =   12
            Top             =   855
            Width           =   5295
         End
         Begin VB.Label lblCorretor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Corretor:"
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
            Left            =   30
            TabIndex        =   9
            Top             =   885
            Width           =   810
         End
         Begin VB.Label lblDataFinal 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   1140
            TabIndex        =   3
            Top             =   15
            Width           =   1695
         End
         Begin VB.Label lblCidade 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cidade:"
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
            Index           =   2
            Left            =   30
            TabIndex        =   17
            Top             =   1725
            Width           =   675
         End
         Begin VB.Label lblEstado 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Estado:"
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
            Index           =   1
            Left            =   45
            TabIndex        =   13
            Top             =   1335
            Width           =   645
         End
         Begin VB.Label lblNomeCliente 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
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
            Left            =   3675
            TabIndex        =   8
            Top             =   405
            Width           =   5295
         End
         Begin VB.Label lblNumeroCPF 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "No. do CPF:"
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
            Index           =   1
            Left            =   30
            TabIndex        =   5
            Top             =   435
            Width           =   1020
         End
         Begin VB.Label lblData 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Data:"
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
            Left            =   45
            TabIndex        =   2
            Top             =   15
            Width           =   480
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
         Left            =   1215
         TabIndex        =   1
         Top             =   15
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lvwTransacao 
      Height          =   360
      Left            =   4350
      TabIndex        =   22
      Top             =   3480
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IdTransacao"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IdCliente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "NomeCliente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "NumeroCPF"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "IdCorretor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "NomeCorretor"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "IdCidade"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Ativo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Data"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvwCliente 
      Height          =   360
      Left            =   1650
      TabIndex        =   26
      Top             =   3465
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IdCliente"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nome"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "NumeroCPF"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image imgIcone 
      Height          =   555
      Left            =   105
      Picture         =   "frmTransacao.frx":0038
      Top             =   75
      Width           =   555
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   17955
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Left            =   -1830
      Picture         =   "frmTransacao.frx":0578
      Top             =   675
      Width           =   10740
   End
End
Attribute VB_Name = "frmTransacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Private vop_TransacaoNegocios As New clsTransacaoNegocios
Private vop_ClientesNegocios As New clsClientesNegocios
Private vop_CorretoresNegocios As New clsCorretoresNegocios
Private vop_System As New clsSystem

'Variaveis de controle do form
Private vbp_Transacao As Boolean                             'Verifica uma inclusao ou alteracao
Private vsl_Estado As String
Private vsl_Cidade As String


'Eventos
Private Sub Form_Activate()
   
   Me.Refresh
   
End Sub

Public Sub Form_Load()
    
    vbp_Transacao = False              'Verifica uma inclusao ou alteracao
    
    'Inicializa entrada e saida
    Call InicializaTransacao
    
    Call ComboBox(cmbLocalizarEstado, "Estados", "ID_Estado", "Nome", Empty)
    If cmbLocalizarEstado.ListCount > 0 Then
       cmbLocalizarEstado.ListIndex = 0
    End If
    
    Call ComboBox(cmbLocalizarCidade, "Cidades", "ID_Cidade", "Nome", "WHERE ID_Estado = " & CStr(cmbLocalizarEstado.ItemData(cmbLocalizarEstado.ListIndex)))
    If cmbLocalizarCidade.ListCount > 0 Then
       cmbLocalizarCidade.ListIndex = 0
    End If
    
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
   
   Set frmTransacao = Nothing
   
End Sub

Private Sub calData_Click()
    lblDataFinal.Caption = calData.Value
End Sub

Private Sub cmbLocalizarEstado_Click()

    Call ComboBox(cmbLocalizarCidade, "Cidades", "ID_Cidade", "Nome", "WHERE ID_Estado = " & CStr(cmbLocalizarEstado.ItemData(cmbLocalizarEstado.ListIndex)))
    If cmbLocalizarCidade.ListCount > 0 Then
       cmbLocalizarCidade.ListIndex = 0
    End If

End Sub

Private Sub fraTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' Converter as coordenadas do mouse de twips para pixels
    Dim mouseX As Single
    Dim mouseY As Single
    
    mouseX = X / Screen.TwipsPerPixelX
    mouseY = Y / Screen.TwipsPerPixelY
    
    'Text2.text = mouseX
    'Text3.text = mouseY
    
    If ((mouseX >= 299 And mouseX <= 355) And (mouseY >= 2 And mouseY <= 20)) Then
       calData.Height = 1890
    Else
       calData.Height = 315
       lblDataFinal.Caption = calData.Value
    End If
    
End Sub

Private Sub cmdExcluir_Click()
    If Trim$(lblCodigo.Caption) = Empty Or Trim$(lblCodigo.Caption) = "0" Then Exit Sub

    If MsgBox("Confirma a Exclusão ?", vbQuestion + vbYesNo, "Confirme !") = vbYes Then
      Set vop_TransacaoNegocios = New clsTransacaoNegocios
          vop_TransacaoNegocios.IdTransacao = lblCodigo.Caption
          If vop_TransacaoNegocios.ExcluirTransacao() = True Then
             MsgBox "Transação excluída com sucesso !", vbExclamation, "Transacao"
             'Atualiza grid
             Call frmMain.CarregarGridTransacao
            
             'Inicializa entrada e saida
             Call InicializaTransacao
            
             'Valida entrada de dados
             If vbp_Transacao = True Then
                Call cmdFechar_Click
             End If
          End If
      Set vop_TransacaoNegocios = Nothing
    End If
End Sub

Private Sub cmdGravar_Click()
Dim vsp_Mensagem As String

   'Valida mensagem
   If vbp_Transacao = False Then
      vsp_Mensagem = "Confirma a Inclusão ?"
   Else
      vsp_Mensagem = "Confirma a Alteração ?"
   End If

   'Valida entrada de dados
   If VerCampos = False Then Exit Sub
   If MsgBox(vsp_Mensagem, vbQuestion + vbYesNo, "Confirme !") = vbYes Then
      
      Set vop_TransacaoNegocios = New clsTransacaoNegocios
          vop_TransacaoNegocios.IdTransacao = lblCodigo.Caption
          vop_TransacaoNegocios.Numero_CPF = msbCPF.text
          vop_TransacaoNegocios.IdCorretor = txtNumeroCorretor.text
          vop_TransacaoNegocios.IdCidade = Int(cmbLocalizarCidade.ItemData(cmbLocalizarCidade.ListIndex))
          vop_TransacaoNegocios.Data_Transacao = lblDataFinal.Caption
          vop_TransacaoNegocios.Ativo = IIf(chkAtivo.Value = 0, 0, 1)
          If vbp_Transacao = False Then
             If vop_TransacaoNegocios.IncluirTransacao(lvwTransacao) = True Then
                MsgBox "Transacao cadastrada com sucesso !", vbExclamation, "Transacao"
             End If
          Else
             If vop_TransacaoNegocios.AlterarTransacao(lvwTransacao) = True Then
                MsgBox "Transacao alterada com sucesso !", vbExclamation, "Transacao"
             End If
          End If
      Set vop_TransacaoNegocios = Nothing
      
   End If
   
   'Atualiza grid
   Call frmMain.CarregarGridTransacao
   
   'Inicializa entrada e saida
   Call InicializaTransacao
   
   'Valida entrada de dados
   If vbp_Transacao = True Then
      Call cmdFechar_Click
   End If
   
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

'CPF
Private Sub msbCPF_Change()

   If VerNumeros = False Then Exit Sub

End Sub

Private Sub msbCPF_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If

End Sub

Private Sub msbCPF_KeyPress(KeyAscii As Integer)

    'Permite Backspace e Enter
    If KeyAscii = vbKeyBack Then Exit Sub

    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If

    'Permite apenas números e formato de Moedas
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0

End Sub

Private Sub msbCPF_LostFocus()

   If VerCliente() = False Then Exit Sub

End Sub

Private Sub cmdPesquisar_Click(Index As Integer)

On Error GoTo TrataErros

   Set vop_System = New clsSystem
       vop_System.Coluna_01 = "No dp cartão:"
       vop_System.Coluna_02 = "Cliente:"
       vop_System.FormataColuna_01 = "000"
       vop_System.TamanhoColuna_01 = "2200"
       If vop_System.FindLista("Clientes", "Numero_CPF", "Nome_Cliente", frmFindLista.lvwLista) = True Then
          frmFindLista.Caption = "Cliente"
          frmFindLista.optSort01.Caption = "No. do cartão"
          frmFindLista.optSort02.Caption = "Cliente"
          frmFindLista.Show vbModal
       Else
          'msbCPF.SetFocus
          Set vop_System = Nothing
          Exit Sub
       End If
    Set vop_System = Nothing

    If Trim(frmFindLista.lvwLista.SelectedItem.text) <> Empty Then
       msbCPF.text = Trim(frmFindLista.lvwLista.SelectedItem)
       lblNomeCliente.Caption = frmFindLista.lvwLista.ListItems(frmFindLista.lvwLista.SelectedItem.Index).SubItems(1)
       Unload frmFindLista

    Else
       Unload frmFindLista
       'msbCPF.SetFocus
    End If

TrataErros:
    If Err.Number <> 0 Then
       MsgBox "Não foi possível listar as Áreas !", vbCritical
       Err.Clear
    End If

End Sub

'Cliente
Private Sub txtNumeroCorretor_Change()
    
    If VerNumeros = False Then Exit Sub
    
End Sub

Private Sub txtNumeroCorretor_GotFocus()
    
    txtNumeroCorretor.SelStart = 0
    txtNumeroCorretor.SelLength = Len(txtNumeroCorretor.text)
    
End Sub

Private Sub txtNumeroCorretor_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyUp Then
        Sendkeys "+{TAB}"
    End If
    If KeyCode = vbKeyDown Then
        Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub txtNumeroCorretor_KeyPress(KeyAscii As Integer)
    
    'Permite Backspace e Enter
    If KeyAscii = vbKeyBack Then Exit Sub
    
    If KeyAscii = 13 Then
        Sendkeys "{TAB}"
        KeyAscii = 0
    End If
    
    'Permite apenas números e formato de Moedas
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
            
End Sub

Private Sub txtNumeroCorretor_LostFocus()

    If VerCorretor() = False Then Exit Sub
    
End Sub

Private Sub cmdPesquisarCorretor_Click()

On Error GoTo TrataErros

   Set vop_System = New clsSystem
       vop_System.Coluna_01 = "No dp cartão:"
       vop_System.Coluna_02 = "Cliente:"
       vop_System.FormataColuna_01 = "000"
       vop_System.TamanhoColuna_01 = "2200"
       If vop_System.FindLista("Corretores", "ID_Corretor", "Nome_Corretor", frmFindLista.lvwLista) = True Then
          frmFindLista.Caption = "Corretor"
          frmFindLista.optSort01.Caption = "Cod. Corretor"
          frmFindLista.optSort02.Caption = "Corretor"
          frmFindLista.Show vbModal
       Else
          'msbCPF.SetFocus
          Set vop_System = Nothing
          Exit Sub
       End If
    Set vop_System = Nothing

    If Trim(frmFindLista.lvwLista.SelectedItem.text) <> Empty Then
       txtNumeroCorretor.text = Trim(frmFindLista.lvwLista.SelectedItem)
       lblNomeCorretor.Caption = frmFindLista.lvwLista.ListItems(frmFindLista.lvwLista.SelectedItem.Index).SubItems(1)
       Unload frmFindLista

    Else
       Unload frmFindLista
       'xtNumeroCorretor.SetFocus
    End If

TrataErros:
    If Err.Number <> 0 Then
       MsgBox "Não foi possível listar as Áreas !", vbCritical
       Err.Clear
    End If

End Sub


'Funcoes
Function Editar(ByVal pIdTransacao As Integer) As Boolean
        
    'Verifica uma inclusao ou alteracao da transacao
    vbp_Transacao = True
    'Controle de exibicao
    lblCodigo.Visible = True
    cmdExcluir.Visible = True
    
    Set vop_TransacaoNegocios = New clsTransacaoNegocios
    
        If vop_TransacaoNegocios.PesquisarTransacao(lvwTransacao, pIdTransacao, 0) = True Then
           lblCodigo.Caption = pIdTransacao
           msbCPF.text = vop_TransacaoNegocios.Numero_CPF
           txtNumeroCorretor.text = vop_TransacaoNegocios.IdCorretor
           chkAtivo.Value = vop_TransacaoNegocios.Ativo
           lblDataFinal.Caption = vop_TransacaoNegocios.Data_Transacao

           'Nome do cliente
           Set vop_ClientesNegocios = New clsClientesNegocios
              If vop_ClientesNegocios.PesquisarCliente(lvwCliente, msbCPF.text, 2) = True Then
                 lvwTransacao.ListItems(lvwTransacao.ListItems.Count).SubItems(1) = lvwCliente.ListItems(lvwCliente.ListItems.Count).text
                 lvwTransacao.ListItems(lvwTransacao.ListItems.Count).SubItems(2) = lvwCliente.ListItems(lvwCliente.ListItems.Count).SubItems(1)
                 lblNomeCliente.Caption = lvwCliente.ListItems(lvwCliente.ListItems.Count).SubItems(1)
              End If
           Set vop_ClientesNegocios = Nothing
           
           'Nome do corretor
           Set vop_CorretoresNegocios = New clsCorretoresNegocios
              If vop_CorretoresNegocios.PesquisarCorretor(lvwCorretor, txtNumeroCorretor.text, 0) = True Then
                 lvwTransacao.ListItems(lvwTransacao.ListItems.Count).SubItems(5) = lvwCorretor.ListItems(lvwCorretor.ListItems.Count).SubItems(1)
                 lblNomeCorretor.Caption = lvwCorretor.ListItems(lvwCorretor.ListItems.Count).SubItems(1)
              End If
           Set vop_ClientesNegocios = Nothing
           
           'Estado e Cidade
           If vop_TransacaoNegocios.PesquisarCidade(vsl_Estado, vsl_Cidade, vop_TransacaoNegocios.IdCidade) = True Then
              If vsl_Estado <> Empty Then
                 cmbLocalizarEstado.text = vsl_Estado
              End If
              If vsl_Cidade <> Empty Then
                 cmbLocalizarCidade.text = vsl_Cidade
              End If
           End If
           
           'Limpa lista
           lvwTransacao.ListItems.Clear
           lvwCliente.ListItems.Clear
        Else
            MsgBox "Não foi possível encontrar a Transação !", vbCritical, "Transação"
        End If
          
    Set vop_TransacaoNegocios = Nothing
    
    Me.Show vbModal

End Function

Private Function VerCliente() As Boolean

    If IsNumeric(RemoverPontosEUnderscores(msbCPF.text)) = False Then
       MsgBox "C.P.F. não existe !", vbCritical, "Transação"
       lblNomeCliente.Caption = Empty
       'msbCPF.SetFocus
       VerCliente = False
       Exit Function
    ElseIf Trim$(msbCPF.text) <> Empty Then
       Set vop_ClientesNegocios = New clsClientesNegocios
          If vop_ClientesNegocios.PesquisarCliente(lvwCliente, msbCPF, 2) = False Then
             MsgBox "Cartão não existe !", vbCritical, "Transação"
             lblNomeCliente.Caption = Empty
             msbCPF.SetFocus
             VerCliente = False
             Exit Function
          End If
       Set vop_ClientesNegocios = Nothing
    Else
       lblNomeCliente.Caption = Empty
    End If

    VerCliente = True

End Function

Private Function VerCorretor() As Boolean

    If Trim$(txtNumeroCorretor.text) = "0" Then
       MsgBox "Códgo do Corretor não existe !", vbCritical, "Transação"
       lblNomeCorretor.Caption = Empty
       'txtNumeroCorretor.SetFocus
       VerCorretor = False
       Exit Function
    ElseIf Trim$(txtNumeroCorretor.text) <> Empty Then
       Set vop_CorretoresNegocios = New clsCorretoresNegocios
          If vop_CorretoresNegocios.PesquisarCorretor(lvwCorretor, txtNumeroCorretor, 0) = False Then
             MsgBox "Corretor não existe !", vbCritical, "Transação"
             lblNomeCorretor.Caption = Empty
             txtNumeroCorretor.SetFocus
             VerCorretor = False
             Exit Function
          End If
       Set vop_CorretoresNegocios = Nothing
    Else
       lblNomeCorretor.Caption = Empty
    End If

    VerCorretor = True

End Function

Private Function VerNumeros() As Boolean

    If IsNumeric(RemoverPontosEUnderscores(msbCPF.text)) = False Then
       'txtTempoDePreparo.SetFocus
       VerNumeros = False
       Exit Function
    End If

    VerNumeros = True

End Function

Private Function DefaultCampos() As BookmarkEnum

    msbCPF.Mask = "###.###.###-##"
    msbCPF.text = "___.___.___-__"
    
    If cmbLocalizarEstado.ListCount > 0 Then
       cmbLocalizarEstado.ListIndex = 0
    End If
    
    If cmbLocalizarCidade.ListCount > 0 Then
       cmbLocalizarCidade.ListIndex = 0
    End If
    
    lblDataFinal.Caption = calData.Value
    
    chkAtivo.Value = 0

End Function

Private Function InicializaTransacao() As Boolean

    'Limpa entrada de dados
    Call LimpaCampos(Me)
    
    'Inicializa entrada de dados
    Call DefaultCampos

End Function

Function VerCampos() As Boolean
    
    If IsNumeric(RemoverPontosEUnderscores(msbCPF.text)) = False Then
       MsgBox "Informe o CPF do Cliente !", vbExclamation, "Transacao"
       'msbCPF.SetFocus
       VerCampos = False
       Exit Function
    End If
    If Trim$(txtNumeroCorretor.text) = Empty Or Trim$(txtNumeroCorretor.text) = "0" Then
       MsgBox "Informe o código do Corretor !", vbExclamation, "Transacao"
        'txtNumeroCorretor.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(cmbLocalizarCidade.text) = Empty Then
       MsgBox "Informe a cidade da Transação !", vbExclamation, "Transacao"
        'cmbLocalizarCidade.SetFocus
        VerCampos = False
        Exit Function
    End If
    If Trim$(cmbLocalizarEstado.text) = Empty Then
       MsgBox "Informe o estado da Transação !", vbExclamation, "Transacao"
        'cmbLocalizarEstado.SetFocus
        VerCampos = False
        Exit Function
    End If
    
    VerCampos = True

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
        If caractereAtual <> "." And caractereAtual <> "_" And caractereAtual <> "-" Then
            resultado = resultado & caractereAtual ' Adiciona o caractere à string de resultado
        End If
    Next i

    ' Retorna a string sem os caracteres indesejad
    RemoverPontosEUnderscores = resultado
End Function


