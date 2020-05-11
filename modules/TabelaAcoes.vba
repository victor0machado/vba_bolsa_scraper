Attribute VB_Name = "TabelaAcoes"
Sub InserirAcao()
  Dim acao, tabela, linha
  Set tabela = ActiveSheet.ListObjects("tblAcoes")
  
  acao = InputBox("Informe o nome da a��o:" & vbNewLine & _
                   "Exemplo: GGBR4.SA para o link https://br.financas.yahoo.com/quote/GGBR4.SA/", _
                   "Inserir nova a��o")
  
  tabela.ListRows.Add
  linha = tabela.ListRows.Count + 1
  
  If acao <> "" Then
    ActiveSheet.Range("A" & linha) = acao
  Else
    MsgBox "Nenhum texto inserido! Ignorando a��o..."
  End If
End Sub

Sub RemoverAcao()
  Dim acao, tabela, linha
  Set tabela = ActiveSheet.ListObjects("tblAcoes")
  
  acao = InputBox("Informe o nome da a��o a ser removida:", _
                  "Remover a��o")
  
  If acao = "" Then
    MsgBox "Nenhum texto inserido! Ignorando a��o..."
    Exit Sub
  End If
  
  For linha = 2 To tabela.ListRows.Count + 1
    If ActiveSheet.Range("A" & linha) = acao Then
      tabela.ListRows(linha - 1).Delete
      MsgBox "A��o " & acao & " removida!"
      Exit Sub
    End If
  Next linha
  
  ' Se chegar nesse ponto � porque n�o encontou a a��o
  MsgBox "A��o " & acao & " n�o foi encontrada!"
End Sub
