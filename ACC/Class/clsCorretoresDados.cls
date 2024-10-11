VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCorretoresDados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Variáveis do ADO
Private vol_Conexao As New clsConexao
Private vol_Command As New ADODB.Command


Public Property Set Conexao(ByVal vData As ADODB.Connection)
    Set cConexao = vData
End Property


'Metodos
Public Function CarregarGridCorretorRS(ByRef pdtgCorretores As DataGrid, ByRef padoCorretores As Adodc, ByVal pcmbLocalizar As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
On Error GoTo TrataErros
     
   CarregarGridCorretorRS = True
   vol_Conexao.AbrirConexao
   
      padoCorretores.ConnectionString = vol_Conexao.cConexao
      padoCorretores.RecordSource = "SELECT ID_Corretor, Nome_Corretor FROM Corretores (NOLOCK) ORDER BY " + IIf(pcmbLocalizar <= 0, "ID_Corretor", "Nome_Corretor") + " "
      padoCorretores.Refresh
      
     ' Conectar o DataGrid ao ADODC
      Set pdtgCorretores.DataSource = padoCorretores
      
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        CarregarGridCorretorRS = False
        MsgBox "Não foi possível carregar informações do Corretor !", vbExclamation
    End If
    
    
End Function

Public Function PesquisarCorretor(ByRef plvwCorretores As ListView, ByVal pCorretor As String, ByVal pBusca As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim i As Integer
On Error GoTo TrataErros
     
   PesquisarCorretor = True
   vol_Conexao.AbrirConexao
      Set vol_Recordset = vol_Conexao.RetornaRs("SELECT ID_Corretor, Nome_Corretor FROM Corretores (NOLOCK) WHERE " + IIf(pBusca = 0, " ID_Corretor = ", " Nome_Corretor = ") + "'" + pCorretor + "'")
      With vol_Recordset
         If .RecordCount <> 0 And .EOF = False Then
            Do Until .EOF
               plvwCorretores.ListItems.Add , , .Fields(0).Value
               plvwCorretores.ListItems(plvwCorretores.ListItems.Count).SubItems(1) = .Fields(1).Value
               .MoveNext
            Loop
         Else
            PesquisarCorretor = False
         End If
      End With
      Set vol_Recordset = Nothing
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        PesquisarCorretor = False
        MsgBox "Não foi possível encontrar informações da Corretor !", vbExclamation
    End If
End Function

Public Function LocalizarCorretor(ByRef pdtgCorretores As DataGrid, ByRef padoCorretores As Adodc, ByVal ptxtLocalizar As String, ByVal pcmbLocalizar As Integer) As Boolean
Dim vol_Recordset As ADODB.Recordset
Dim i As Integer
On Error GoTo TrataErros
     
   LocalizarCorretor = True
   vol_Conexao.AbrirConexao
   
      padoCorretores.ConnectionString = vol_Conexao.cConexao
      If Trim$(ptxtLocalizar) = Empty Then
         padoCorretores.RecordSource = "SELECT ID_Corretor, Nome_Corretor FROM Corretores (NOLOCK) ORDER BY " + IIf(pcmbLocalizar = 0, "ID_Corretor", "Nome_Corretor") + " "
      Else
         padoCorretores.RecordSource = "SELECT ID_Corretor, Nome_Corretor FROM Corretores (NOLOCK) WHERE " + IIf(pcmbLocalizar = 0, "ID_Corretor", "Nome_Corretor") + " LIKE '%" + Trim$(ptxtLocalizar) + "%' ORDER BY " + IIf(pcmbLocalizar = 0, "ID_Corretor", "Nome_Corretor") + " "
      End If
      padoCorretores.Refresh
      
     ' Conectar o DataGrid ao ADODC
      Set pdtgCorretores.DataSource = padoCorretores
      
   vol_Conexao.FecharConexao
    
TrataErros:
    If Err.Number <> 0 Then
        vol_Conexao.FecharConexao
        Err.Clear
        LocalizarCorretor = False
        MsgBox "Não foi possível encontrar informações da Corretor !", vbExclamation
    End If
End Function

Public Function IncluirCorretor(ByRef plvwCorretor As ListView) As Boolean
Dim Incluir As Boolean
Dim Item As ListItem

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
         
         IncluirCorretor = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
   
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "sp_InserirCorretor"
         vol_Command.Parameters.Refresh
      
         ' Pega o primeiro (e único) item do ListView
         Set Item = plvwCorretor.ListItems(1)
      
         With vol_Command.Parameters
           .Item(1).Value = Item.SubItems(1)  'Nome
         End With
   
         vol_Command.Execute
         
      vol_Conexao.cConexao.CommitTrans
   
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       IncluirCorretor = False
       MsgBox "Não foi possível incluir informações da Corretor !", vbExclamation
    End If
End Function

Public Function AlterarCorretor(ByRef plvwCorretor As ListView) As Boolean
Dim Incluir As Boolean
Dim Item As ListItem

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
      
         AlterarCorretor = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "sp_AtualizarCorretor"
         vol_Command.Parameters.Refresh
      
         ' Pega o primeiro (e único) item do ListView
         Set Item = plvwCorretor.ListItems(1)

         With vol_Command.Parameters
           .Item(1).Value = Item.text        'IdCorretor
           .Item(2).Value = Item.SubItems(1) 'Descricao
         End With
      
         vol_Command.Execute
      
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       AlterarCorretor = False
       MsgBox "Não foi possível incluir informações da Corretor !", vbExclamation
    End If
End Function

Public Function ExcluirCorretor(ByVal pIdCorretor As Integer) As Boolean
Dim Incluir As Boolean

On Error GoTo TrataErros
   
   vol_Conexao.AbrirConexao
      vol_Conexao.cConexao.BeginTrans
      
         ExcluirCorretor = True
         Set vol_Command = New ADODB.Command
         Set vol_Command.ActiveConnection = vol_Conexao.cConexao
      
         vol_Command.CommandType = adCmdStoredProc
         vol_Command.CommandText = "sp_ExcluirCorretor"
         vol_Command.Parameters.Refresh
      
         With vol_Command.Parameters
           .Item(1).Value = pIdCorretor
         End With
      
         vol_Command.Execute
   
      vol_Conexao.cConexao.CommitTrans
   vol_Conexao.FecharConexao
   
TrataErros:
    If Err.Number <> 0 Then
       vol_Conexao.cConexao.RollbackTrans
       vol_Conexao.FecharConexao
       Err.Clear
       ExcluirCorretor = False
       MsgBox "Não foi possível excluir informações da Corretor !", vbExclamation
    End If
End Function
























