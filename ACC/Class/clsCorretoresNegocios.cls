VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCorretoresNegocios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Inicializa variáveis da classe
Private vop_CorretoresDados As New clsCorretoresDados      'Variável de acesso classe clsCorretorDados

'Variáveis da classe
Private vsp_Query As String                'Variável usada nas consultas sql
Private vip_Count As Integer               'Variável utilizadas nos loops
Private vbp_RollBack As Boolean            'Define que uma conexão está aberta e necessita de um rollback

'Variáveis das Propriedades
Private mvarIdCorretor As Long              'Identificador da Corretores
Private mvarNome As String                 'Descrição da Corretores




'Propriedades
'IdCorretor
Public Property Let IdCorretor(ByVal vData As Long)
    mvarIdCorretor = vData
End Property

Public Property Get IdCorretor() As Long
    IdCorretor = mvarIdCorretor
End Property

'Nome
Public Property Let Nome(ByVal vData As String)
    mvarNome = vData
End Property

Public Property Get Nome() As String
    Nome = mvarNome
End Property




'Metodos
Public Function PesquisarCorretor(ByRef plvwCorretores As ListView, ByVal pCorretor As String, ByVal pBusca As Integer) As Boolean

    PesquisarCorretor = False
    Screen.MousePointer = vbHourglass
      If vop_CorretoresDados.PesquisarCorretor(plvwCorretores, pCorretor, pBusca) = True Then
         With plvwCorretores
            IdCorretor = Trim$(.ListItems(plvwCorretores.ListItems.Count).text)
            Nome = Trim$(.ListItems(plvwCorretores.ListItems.Count).SubItems(1))
         End With
         PesquisarCorretor = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function LocalizarCorretor(ByRef pdtgCorretores As DataGrid, ByRef padoCorretores As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean

    LocalizarCorretor = False
    Screen.MousePointer = vbHourglass
      If vop_CorretoresDados.LocalizarCorretor(pdtgCorretores, padoCorretores, ptxtLocalizar, pcmbLocalizar) = True Then
         LocalizarCorretor = True
      End If
    Screen.MousePointer = vbDefault
    
End Function

Public Function CarregarGridCorretorRS(ByRef pdtgCorretores As DataGrid, ByRef padoCorretores As Adodc, ByVal pcmbLocalizar As Integer) As Boolean
     
    Screen.MousePointer = vbHourglass
      CarregarGridCorretorRS = vop_CorretoresDados.CarregarGridCorretorRS(pdtgCorretores, padoCorretores, pcmbLocalizar)
    Screen.MousePointer = vbDefault
    
End Function

Public Function IncluirCorretor(ByRef plvwCorretor As ListView) As Boolean
Dim ListItem As ListItem

On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      ' Para adicionar dados na lista Corretor
      Set ListItem = plvwCorretor.ListItems.Add(, , IdCorretor)
      ListItem.SubItems(1) = Nome
      '
      IncluirCorretor = vop_CorretoresDados.IncluirCorretor(plvwCorretor)
      'Limpa lista
      plvwCorretor.ListItems.Clear
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function AlterarCorretor(ByRef plvwCorretor As ListView) As Boolean
Dim ListItem As ListItem

On Error GoTo TrataErros

   Screen.MousePointer = vbHourglass
      'Limpa lista
      plvwCorretor.ListItems.Clear
      ' Para adicionar dados na lista Corretor
      Set ListItem = plvwCorretor.ListItems.Add(, , IdCorretor)
      ListItem.SubItems(1) = Nome
      '
      AlterarCorretor = vop_CorretoresDados.AlterarCorretor(plvwCorretor)
      'Limpa lista
      plvwCorretor.ListItems.Clear
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
    
End Function

Public Function ExcluirCorretor() As Boolean
   Screen.MousePointer = vbHourglass
      ExcluirCorretor = vop_CorretoresDados.ExcluirCorretor(IdCorretor)
   Screen.MousePointer = vbDefault
    
TrataErros:
    If Err.Number <> 0 Then
       Err.Clear
       Screen.MousePointer = vbDefault
    End If
End Function
















