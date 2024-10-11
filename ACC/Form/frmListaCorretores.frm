VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Begin VB.Form frmListaCorretores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Corretores"
   ClientHeight    =   5190
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6915
   LinkTopic       =   "frmListaCorretores"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5190
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdFechar 
      Caption         =   "&Fechar"
      Height          =   360
      Left            =   5745
      TabIndex        =   8
      Top             =   4740
      Width           =   1005
   End
   Begin VB.CommandButton cmdExcluir 
      Caption         =   "&Excluir"
      Height          =   360
      Left            =   4680
      TabIndex        =   7
      Top             =   4740
      Width           =   1005
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Height          =   360
      Left            =   3630
      TabIndex        =   6
      Top             =   4740
      Width           =   1005
   End
   Begin VB.ComboBox cmbLocalizar 
      Height          =   315
      ItemData        =   "frmListaCorretores.frx":0000
      Left            =   4320
      List            =   "frmListaCorretores.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   825
      Width           =   1410
   End
   Begin VB.TextBox txtLocalizar 
      Height          =   315
      Left            =   900
      MaxLength       =   40
      TabIndex        =   2
      Top             =   825
      Width           =   2475
   End
   Begin MSDataGridLib.DataGrid dtgCorretores 
      Bindings        =   "frmListaCorretores.frx":001C
      Height          =   3345
      Left            =   150
      TabIndex        =   5
      Top             =   1230
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   5900
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "ID_Corretor"
         Caption         =   "   Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "Nome_Corretor"
         Caption         =   "Nome"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   915,024
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   5070,047
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoCorretores 
      Height          =   330
      Left            =   450
      Top             =   4725
      Visible         =   0   'False
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=BDAdministradoraCC"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=BDAdministradoraCC"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT ID_Corretor, Nome_Corretor FROM Corretores ORDER BY ID_Corretor"
      Caption         =   "adoCorretores"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image Image2 
      Height          =   480
      Index           =   12
      Left            =   75
      Picture         =   "frmListaCorretores.frx":0038
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      Height          =   195
      Index           =   2
      Left            =   3780
      TabIndex        =   3
      Top             =   885
      Width           =   360
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Localizar:"
      ForeColor       =   &H00000000&
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   1
      Top             =   900
      Width           =   675
   End
   Begin VB.Label Label23 
      BackColor       =   &H00FFFFFF&
      Height          =   645
      Index           =   11
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   17955
   End
   Begin VB.Image imgLinha 
      Height          =   45
      Index           =   10
      Left            =   -1200
      Picture         =   "frmListaCorretores.frx":0902
      Top             =   660
      Width           =   10740
   End
End
Attribute VB_Name = "frmListaCorretores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variável de acesso as classes
Dim vop_CorretoresNegocios As New clsCorretoresNegocios

'Variaveis de controle do form
Private vil_IdCorretor As Long                               'Identificador da Corretor
Private bMovingProgrammatically As Boolean                  ' Variável de controle para evitar movimentação dupla
Private vil_LastBookmark As Variant                         ' Variável de módulo para armazenar o último Bookmark



'Eventos
Private Sub Form_Activate()
   
    Me.Refresh
    
End Sub

Private Sub Form_Load()
    
    'Carrega Corretores
    Call CarregarGrid
    cmbLocalizar.ListIndex = 0
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

On Error GoTo TrataErros

    'Tecla de atalho da calculadora
    If KeyCode = vbKeyF7 Then
        KeyCode = 0
        Exit Sub
    End If
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
   
   Set frmListaCorretores = Nothing
   
End Sub

Private Sub cmdExcluir_Click()
    If vil_IdCorretor = 0 Then Exit Sub

    If MsgBox("Confirma a Exclusão ?", vbQuestion + vbYesNo, "Confirme !") = vbYes Then
      Set vop_CorretoresNegocios = New clsCorretoresNegocios
          vop_CorretoresNegocios.IdCorretor = vil_IdCorretor
          If vop_CorretoresNegocios.ExcluirCorretor() = True Then
             MsgBox "Corretor excluído com sucesso !", vbExclamation, "Corretor"
             'Limpa variável
             txtLocalizar.text = Empty
             Call CarregarGrid
          End If
      Set vop_CorretoresNegocios = Nothing
    End If
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdNovo_Click()

On Error GoTo TrataErros
   
    'Inicia form
    frmCorretor.Show vbModal
    
TrataErros:
    If Err.Number <> 0 Then
       Set vop_CorretoresNegocios = Nothing
       Exit Sub
    End If
   
End Sub


Private Sub cmbLocalizar_Click()

On Error GoTo TrataErros

   Set vop_CorretoresNegocios = New clsCorretoresNegocios
       Call vop_CorretoresNegocios.LocalizarCorretor(dtgCorretores, adoCorretores, txtLocalizar.text, cmbLocalizar.ListIndex)
   Set vop_CorretoresNegocios = Nothing

   txtLocalizar.text = Empty

TrataErros:
    If Err.Number <> 0 Then
       Set vop_CorretoresNegocios = Nothing
       Exit Sub
    End If

End Sub

Private Sub dtgCorretores_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    ' Só atualiza se não estiver movendo programaticamente
    If Not bMovingProgrammatically Then
        If dtgCorretores.Bookmark > 0 Then
            If dtgCorretores.SelBookmarks.Count > 0 Then
                dtgCorretores.SelBookmarks.Remove 0
            End If
            dtgCorretores.SelBookmarks.Add dtgCorretores.Bookmark
            If IsNull(dtgCorretores.Columns(0).text) Or dtgCorretores.Columns(0).text = "" Then
                vil_IdCorretor = 0
            Else
                vil_IdCorretor = CInt(dtgCorretores.Columns(0).text)
            End If
        End If
    End If
End Sub

Private Sub dtgCorretores_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyDown
            MoveRecord adoCorretores.Recordset, False
            'Valida fim do DataGrid
        Case vbKeyUp
            MoveRecord adoCorretores.Recordset, True
    End Select
End Sub

Private Sub MoveRecord(rs As ADODB.Recordset, MovePrevious As Boolean)
    If rs.RecordCount > 0 Then
        bMovingProgrammatically = True
        
        If dtgCorretores.SelBookmarks.Count > 0 Then
            dtgCorretores.SelBookmarks.Remove 0
        End If
        
        If MovePrevious Then
            rs.MovePrevious
            If rs.BOF Then
                rs.MoveFirst
                dtgCorretores.SelBookmarks.Add dtgCorretores.Bookmark
            Else
                rs.MoveNext
            End If
        Else
            rs.MoveNext
            If rs.EOF Then
                rs.MoveLast
                dtgCorretores.SelBookmarks.Add dtgCorretores.Bookmark
            Else
               rs.MovePrevious
            End If
        End If
        
        dtgCorretores.Bookmark = rs.Bookmark
        
        bMovingProgrammatically = False ' Reseta o flag
    End If
End Sub

Private Sub dtgCorretores_DblClick()
Dim vvl_BookMark As Variant
Dim vil_RowIndex As Long

On Error GoTo TrataErros

    vil_RowIndex = dtgCorretores.Row
    vvl_BookMark = dtgCorretores.RowBookmark(vil_RowIndex)
    If vvl_BookMark = Empty Then Exit Sub

    Call frmCorretor.Form_Load
    Call frmCorretor.Editar(vil_IdCorretor)

    dtgCorretores.Bookmark = vvl_BookMark
    dtgCorretores.Scroll 0, dtgCorretores.RowContaining(vvl_BookMark)
    dtgCorretores.SelBookmarks.Add vvl_BookMark
    dtgCorretores.Refresh

TrataErros:
    If Err.Number <> 0 Then Exit Sub

End Sub

Private Sub txtLocalizar_Change()

On Error GoTo TrataErros

   Set vop_CorretoresNegocios = New clsCorretoresNegocios
       Call vop_CorretoresNegocios.LocalizarCorretor(dtgCorretores, adoCorretores, txtLocalizar.text, cmbLocalizar.ListIndex)
   Set vop_CorretoresNegocios = Nothing

TrataErros:
    If Err.Number <> 0 Then
       Set vop_CorretoresNegocios = Nothing
       Exit Sub
    End If

End Sub

'Medotos
Public Sub CarregarGrid()
Dim vbl_Carregar As Boolean '

On Error GoTo TrataErros '

    Set vop_CorretoresNegocios = New clsCorretoresNegocios '

        vbl_Carregar = vop_CorretoresNegocios.CarregarGridCorretorRS(dtgCorretores, adoCorretores, cmbLocalizar.ListIndex)
        If vbl_Carregar = True And adoCorretores.Recordset.RecordCount > 0 Then
           cmdNovo.Left = 3630
           cmdExcluir.Left = 4680
           cmdExcluir.Visible = True
        Else
           dtgCorretores.Refresh
           cmdNovo.Left = 4680
           cmdExcluir.Visible = False
        End If

    Set vop_CorretoresNegocios = Nothing

TrataErros:
    If Err.Number <> 0 Then
       Set vop_CorretoresNegocios = Nothing
       Exit Sub
    End If

End Sub

