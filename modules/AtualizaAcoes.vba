Attribute VB_Name = "AtualizaAcoes"
Sub ImportaTabelaHTML(url, planilha)
  Dim cont_html, linha, coluna, tabela, tabela_html
  Dim coluna_inicio

  Set cont_html = CreateObject("htmlfile")
  
  With CreateObject("msxml2.xmlhttp")
    .Open "GET", url, False
    .send
    cont_html.Body.innerHTML = .responseText
  End With
  
  Set tabela_html = cont_html.getElementsByTagName("table")(0)
  Set tabela = planilha.ListObjects("tblDados")

  tabela.ListRows.Add
  linha = tabela.ListRows.Count + 1

  With tabela_html
    Debug.Print (.innerHTML)
    planilha.Range("A" & linha) = Date
    planilha.Range("B" & linha) = Time
    planilha.Range("C" & linha) = .Rows(0).Cells(1).innerText
    planilha.Range("D" & linha) = .Rows(2).Cells(1).innerText
    planilha.Range("E" & linha) = .Rows(3).Cells(1).innerText
  End With
End Sub

Sub CriaPlanilha(acao)
  Dim nova_planilha, tabela
  
  Set nova_planilha = Sheets.Add(After:=ActiveSheet)
  nova_planilha.name = acao
  
  Set tabela = nova_planilha.ListObjects.Add(xlSrcRange, Range("$A$1:$E$1"), , xlYes)
  
  Range("A1") = "Data"
  Range("B1") = "Hora"
  Range("C1") = "Fechamento Anterior"
  Range("D1") = "Valor Compra"
  Range("E1") = "Valor Venda"
  
  tabela.name = "tblDados"
  
  Worksheets("Home").Select
End Sub

Function PlanilhaJaExiste(acao)
  For Each ws In Worksheets
    If acao = ws.name Then
      PlanilhaJaExiste = True
      Exit Function
    End If
  Next ws
  
  PlanilhaJaExiste = False
End Function

Sub AtualizaPlanilhas()
  Dim acao, tabela, linha, url_prefix
  
  Set tabela = ActiveSheet.ListObjects("tblAcoes")
  url_prefix = "https://br.financas.yahoo.com/quote/"
  
  For linha = 2 To tabela.ListRows.Count + 1
    acao = ActiveSheet.Range("A" & linha)
    If PlanilhaJaExiste(acao) = False Then
      CriaPlanilha acao
    End If
    
    ImportaTabelaHTML url_prefix & acao, Worksheets(acao)
  Next linha
  
  MsgBox ("Concluído!")
End Sub
