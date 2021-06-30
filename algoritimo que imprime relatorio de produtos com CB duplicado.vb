
'Teste Bruno
Private Sub RemoverDuplicidade()

    'Define as variáveis
    Dim SQL
    Dim rsProdutosDuplicados As New Recordset
    Dim rsProdutosComMesmoCB As New Recordset
    Dim rsAtualizarProduto As New Recordset
    
    SQL = "SELECT codigobarrasprincipal," & vbNewLine
    SQL = SQL & "      Count(*) as 'repetcoes' " & vbNewLine
    SQL = SQL & "FROM ProdutosItens" & vbNewLine
    SQL = SQL & "WHERE CodigoBarrasPrincipal <> ''" & vbNewLine
    SQL = SQL & "GROUP BY codigobarrasprincipal" & vbNewLine
    SQL = SQL & "HAVING Count(*) > 1 "
    
    If rsProdutosDuplicados.State = 1 Then rsProdutosDuplicados.Close
    rsProdutosDuplicados.Open SQL, SAGEDLL.CNBDADOS, adOpenForwardOnly, adLockReadOnly
    SQL = ""
    
    ' 1° Nivel - Verifica se existem produtos com código de barras Duplicados
    Do While Not rsProdutosDuplicados.EOF

        SQL = "Select Produtos.Codigo AS Codigo, ProdutosItens.CodigoBarrasPrincipal," & vbNewLine
        SQL = SQL & "Produtos.Situacao as Situacao, ProdutosItens.EstoqueInterno, Produtos.Nome as Nome" & vbNewLine
        SQL = SQL & "From Produtos" & vbNewLine
        SQL = SQL & "INNER JOIN ProdutosItens ON Produtos.Codigo = ProdutosItens.CodigoProduto" & vbNewLine
        SQL = SQL & "where CodigoBarrasPrincipal = '" & rsProdutosDuplicados("codigobarrasprincipal") & "'"
        
        If rsProdutosComMesmoCB.State = 1 Then rsProdutosComMesmoCB.Close
        rsProdutosComMesmoCB.Open SQL, SAGEDLL.CNBDADOS, adOpenForwardOnly, adLockOptimistic
        
        ' 2° Nivel - Verifica se existem produtos com código de barras Duplicados
        Dim i As Integer
        For i = 1 To Val(rsProdutosDuplicados("repetcoes"))
            'rsProdutosComMesmoCB("CodigoBarrasPrincipal") = ""
            'rsProdutosComMesmoCB.MoveNext
            ImprimeProdutosDuplicados rsProdutosComMesmoCB("Codigo"), rsProdutosComMesmoCB("Nome"), rsProdutosComMesmoCB("Situacao"), rsProdutosComMesmoCB("estoqueInterno"), rsProdutosDuplicados("codigobarrasprincipal")
            
            rsProdutosComMesmoCB.MoveNext
        Next

        rsProdutosDuplicados.MoveNext
    Loop
    
    'No Sage utiliza esse método para finalizar o recordset
    RecordsetFinaliza rsProdutosDuplicados
    RecordsetFinaliza rsProdutosComMesmoCB
    
End Sub
'FIM teste Bruno


Public Function ImprimeProdutosDuplicados(CODIGO As String, NOME As String, SITUACAO As String, ESTOQUEinterno As String, CODIGObarras As String)

    Dim INDICE As Integer
    Dim ARQCONTEUDO() As String
    Dim ARQNome As String, ARQPasta As String, SQL As String
    Dim rsEntradasSaidas As New Recordset
    
    'NOME LOG + ENTRADA OU SAIDA + CÓDIGO + .log
    ARQNome = "CB" & CODIGObarras & ".TXT"
    
    'ARQPasta = App.Path & "\LOG"
    ARQPasta = "C:\Users\brunomss\Desktop" & "\LOG"
    
    'Cria pasta se não existir
    If FSO.FolderExists(ARQPasta) = False Then FSO.CreateFolder (ARQPasta)
    
    Close #99
    
    'Altera o arquivo caso ele exista
    If FSO.FileExists(ARQPasta & "\" & ARQNome) = True Then
        Open (ARQPasta & "\" & ARQNome) For Append As #99
    Else
        'Cria o arquivo
        Open (ARQPasta & "\" & ARQNome) For Output As #99
        
        INDICE = 0
        ReDim Preserve ARQCONTEUDO(INDICE)
    
        ARQCONTEUDO(INDICE) = "*************************************************"
        Print #99, ARQCONTEUDO(INDICE)
        ARQCONTEUDO(INDICE) = "CÓDIGO DE BARRAS: " & CODIGObarras
        Print #99, ARQCONTEUDO(INDICE)
        ARQCONTEUDO(INDICE) = "*************************************************"
        Print #99, ARQCONTEUDO(INDICE)
        ARQCONTEUDO(INDICE) = " "
        Print #99, ARQCONTEUDO(INDICE)
    End If
        
    
    '***********************************************************
    '****************IMPRIME O ARQUIVO DE LOG*******************
    '***********************************************************
    
    'Redeclara a Variavel e Limpa o Conteudo
    INDICE = 0
    ReDim Preserve ARQCONTEUDO(INDICE)
    
    
    'DATA DO EVENTO
    ARQCONTEUDO(INDICE) = " Código.............: " & CODIGO
    Print #99, ARQCONTEUDO(INDICE)
    'ORIGEM
    ARQCONTEUDO(INDICE) = " Nome...............: " & NOME
    Print #99, ARQCONTEUDO(INDICE)
    'ORIGEM CODIGO
    ARQCONTEUDO(INDICE) = " Situacão...........: " & SITUACAO
    Print #99, ARQCONTEUDO(INDICE)
    'NUMERO
    ARQCONTEUDO(INDICE) = " Estoque Interno....: " & ESTOQUEinterno
    Print #99, ARQCONTEUDO(INDICE)
    'LOTE
    ARQCONTEUDO(INDICE) = " Últimas 3 Saídas...: "
    Print #99, ARQCONTEUDO(INDICE)
    SQL = "SELECT top 3 CodigoSaida, Saidas.Data" & vbNewLine
    SQL = SQL & "FROM SaidasProdutos" & vbNewLine
    SQL = SQL & "INNER JOIN Saidas on SaidasProdutos.CodigoSaida = Saidas.Codigo" & vbNewLine
    SQL = SQL & "WHERE SaidasProdutos.CodigoProduto =" & CODIGO & vbNewLine
    SQL = SQL & "ORDER BY Saidas.Data DESC"
    
    If rsEntradasSaidas.State = 1 Then rsEntradasSaidas.Close
    rsEntradasSaidas.Open SQL, SAGEDLL.CNBDADOS, adOpenForwardOnly, adLockOptimistic
    Do While Not rsEntradasSaidas.EOF
        ARQCONTEUDO(INDICE) = "     - Data.: " & rsEntradasSaidas("Data") & " - (Cód:" & rsEntradasSaidas("CodigoSaida") & ")"
        Print #99, ARQCONTEUDO(INDICE)
        rsEntradasSaidas.MoveNext
    Loop
    
    'CHAVE DE ACESSO
    ARQCONTEUDO(INDICE) = " Últimas 3 Entradas.: "
    Print #99, ARQCONTEUDO(INDICE)
    
    SQL = "SELECT top 3 CodigoEntrada, Entradas.Data" & vbNewLine
    SQL = SQL & "FROM EntradasProdutos" & vbNewLine
    SQL = SQL & "INNER JOIN Entradas on EntradasProdutos.CodigoEntrada = Entradas.Codigo" & vbNewLine
    SQL = SQL & "WHERE EntradasProdutos.CodigoProduto = " & CODIGO & vbNewLine
    SQL = SQL & "ORDER BY Entradas.Data DESC"
    
    If rsEntradasSaidas.State = 1 Then rsEntradasSaidas.Close
    rsEntradasSaidas.Open SQL, SAGEDLL.CNBDADOS, adOpenForwardOnly, adLockOptimistic
    Do While Not rsEntradasSaidas.EOF
        ARQCONTEUDO(INDICE) = "     - Data.: " & rsEntradasSaidas("Data") & " - (Cód:" & rsEntradasSaidas("CodigoEntrada") & ")"
        Print #99, ARQCONTEUDO(INDICE)
        rsEntradasSaidas.MoveNext
    Loop
    
    'SEPARADOR
    ARQCONTEUDO(INDICE) = "---------------------------------------"
    Print #99, ARQCONTEUDO(INDICE)
    
    Close #99
    
    RecordsetFinaliza rsEntradasSaidas

End Function