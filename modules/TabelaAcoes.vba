Attribute VB_Name = "TabelaAcoes"
Sub InserirAcao()
  Dim acao, tabela, linha
  Set tabela = ActiveSheet.ListObjects("tblAcoes")
  
  acao = InputBox("Informe o nome da ação:" & vbNewLine & _
                   "Exemplo: GGBR4.SA para o link https://br.financas.yahoo.com/quote/GGBR4.SA/", _
                   "Inserir nova ação")
  
  tabela.ListRows.Add
  linha = tabela.ListRows.Count + 1
  
  If acao <> "" Then
    ActiveSheet.Range("A" & linha) = acao
  Else
    MsgBox "Nenhum texto inserido! Ignorando ação..."
  End If
End Sub

Sub RemoverAcao()
  Dim acao, tabela, linha
  Set tabela = ActiveSheet.ListObjects("tblAcoes")
  
  acao = InputBox("Informe o nome da ação a ser removida:", _
                  "Remover ação")
  
  If acao = "" Then
    MsgBox "Nenhum texto inserido! Ignorando ação..."
    Exit Sub
  End If
  
  For linha = 2 To tabela.ListRows.Count + 1
    If ActiveSheet.Range("A" & linha) = acao Then
      tabela.ListRows(linha - 1).Delete
      MsgBox "Ação " & acao & " removida!"
      Exit Sub
    End If
  Next linha
  
  ' Se chegar nesse ponto é porque não encontou a ação
  MsgBox "Ação " & acao & " não foi encontrada!"
End Sub
