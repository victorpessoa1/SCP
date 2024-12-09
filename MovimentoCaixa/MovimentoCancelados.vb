Private Sub cmdCancelados_Click()
On Error GoTo Erro
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim RSMovCaixa As ADODB.Recordset
    Set RSMovCaixa = New ADODB.Recordset
    Dim TransacaoAtiva As Boolean
    TransacaoAtiva = False

    ' Limpa a tabela tblMovimentoCaixa
    DB.Execute "Delete From tblMovimentoCaixa"
    RSMovCaixa.Open "Select * From tblMovimentoCaixa", DB, adOpenDynamic
    
    ' Verifica se o relatório está em uso
    If RSMovCaixa.RecordCount > 0 Then
        nMC = MsgBox("O Relatório está em uso. Liberar?", vbQuestion + vbYesNo)
        If nMC = vbYes Then
            DB.BeginTrans
            TransacaoAtiva = True
            DB.Execute "Delete From tblMovimentoCaixa"
            DB.CommitTrans
            TransacaoAtiva = False
        Else
            RSMovCaixa.Close
            Exit Sub
        End If
    End If
    RSMovCaixa.Close

    ' Desabilita o botão cmdCancelados e exibe o processo de carregamento
    Me.cmdCancelados.Enabled = False
    Me.cmdProcessando.Visible = True
    Me.Refresh

    ' Abrindo o Recordset para obter os dados de tblFinanceiro e tblTitulo
    RS.Open "SELECT t.*, f.*, t.Livro as Livro " & _
    "FROM tblFinanceiro AS f " & _
    "LEFT JOIN tblTitulo AS t ON t.Protocolo_Cartorio = f.Protocolo " & _
    "WHERE (f.Tipo_Ocorrencia = 'a' AND t.Anulado = '0' AND t.CancelaBanco = '0'  AND (Baixa_Lote = 0 OR Baixa_Lote IS NULL) AND f.Data_Cancelado BETWEEN '" & Format(Me.txtDataInicio, "yyyy/mm/dd") & "' AND '" & Format(Me.txtDataFim, "yyyy/mm/dd") & "' AND f.Custas <> 0) " & _
    "OR (f.Tipo_Ocorrencia = 'a' AND t.CancelaBanco = 1 AND f.Rec_Canc = 1 AND Baixa_Lote = 1 AND f.Custas <> 0 AND f.Rec_Data BETWEEN '" & Format(Me.txtDataInicio, "yyyy/mm/dd") & "' AND '" & Format(Me.txtDataFim, "yyyy/mm/dd") & "')" & _
    "ORDER BY f.Protocolo", DB, adOpenDynamic


    ' Verifica se há registros retornados

    If RS.RecordCount = 0 Then
        Me.cmdProcessando.Visible = False
        Me.Refresh
        MsgBox "Sem Dados para exibir!", vbInformation
        RS.Close
        Exit Sub
    End If

    ' Iniciar transação para processar todos os registros
    DB.BeginTrans
    TransacaoAtiva = True

    ' Confirma a impressão
    nResp = vbYes ' ou utilize o MsgBox se desejar perguntar ao usuário
    If nResp = vbYes Then
        ' Loop pelos registros retornados
        Do While Not RS.EOF
            If IsNull(RS!Estorno) Or RS!Estorno = False Then
                    ' Tratamento de nulos e valores para inserção
    TxCartao = IIf(IsNull(RS!TaxaCartao), 0, RS!TaxaCartao)
    
    ' Verifique se TxCartao é numérico antes de realizar a conversão
    If IsNumeric(TxCartao) Then
        TxCartao = Replace(TxCartao, ",", ".")
    Else
        TxCartao = 0 ' Atribua um valor padrão se não for numérico
    End If
    
    ' Verifique se RS!Codigo é 1 e defina TxBanco adequadamente
    If IsNumeric(RS!Codigo) And RS!Codigo = 1 Then
        TxBanco = 1
    Else
        TxBanco = TxCartao
    End If
    
    DataOriginal = RS!Vencimento
    Dim dia As String
    Dim mes As String
    Dim ano As String

    dia = Left(DataOriginal, 2)    ' Pega os dois primeiros caracteres (dia)
    mes = Mid(DataOriginal, 3, 2)  ' Pega os dois caracteres do meio (mês)
    ano = Right(DataOriginal, 4)   ' Pega os quatro últimos caracteres (ano)

    ' Monta a data no formato dd/mm/yyyy
    VencimentoFormatado = mes & "/" & dia & "/" & ano
    
    

    FRJ = IIf(IsNull(RS!FRJ), 0, RS!FRJ)
    FRC = IIf(IsNull(RS!FRC), 0, RS!FRC)
    Cartao = IIf(IsNull(RS!Cartao) Or RS!Cartao = False, 0, 1)
  


    

                ' Insere dados na tblMovimentoCaixa
                If RS!CancelaBanco = True Then
                    DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo, Num_Titulo, Especie_Tit, Vencimento, Data_Apresenta, Origem, Saldo, Custas, Selo, Distrib, Devedor, Sacador, Nosso_Num, Portador, Livro, Folha, Usuario, Texto, TxBanco, ISS, Data_Protestado, FRJ, FRC, Cartao) " & _
                                "Values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & VencimentoFormatado & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RS!Custas, ",", ".") & "','" & Replace(RS!Valor_Selo, ",", ".") & "','" & Replace(RS!Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "','" & RS!Livro & "','" & RS!Pagina & "','" & RS!Usuario & "', 'BANCO', '" & TxBanco & "','" & Replace(RS!ISS, ",", ".") & "', '" & Format(RS!data_protestado, "mm/dd/yyyy") & "', '" & Replace(FRJ, ",", ".") & "', '" & Replace(FRC, ",", ".") & "', " & Cartao & ")")
                Else
                    DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo, Num_Titulo, Especie_Tit, Vencimento, Data_Apresenta, Origem, Saldo, Custas, Selo, Distrib, Devedor, Sacador, Nosso_Num, Portador, Livro, Folha, Usuario, TxBanco, ISS, Data_Protestado, FRJ, FRC, Cartao) " & _
                                "Values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & VencimentoFormatado & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RS!Custas, ",", ".") & "','" & Replace(RS!Valor_Selo, ",", ".") & "','" & Replace(RS!Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "','" & RS!Livro & "','" & RS!Pagina & "','" & RS!Usuario & "', '" & TxBanco & "', '" & Replace(RS!ISS, ",", ".") & "', '" & Format(RS!data_protestado, "mm/dd/yyyy") & "', '" & Replace(FRJ, ",", ".") & "', '" & Replace(FRC, ",", ".") & "', " & Cartao & ")")
                End If
            End If
            RS.MoveNext
        Loop

        ' Se tudo ocorreu bem, confirmar a transação
        DB.CommitTrans
        TransacaoAtiva = False
    Else
        Me.cmdProcessando.Visible = False
        MsgBox "Impressão Cancelada!", vbInformation
        RS.Close
        DB.Execute "Delete From tblMovimentoCaixa"
        Exit Sub
    End If

    ' <<< IMPRESSÃO >>>
    Me.cmdProcessando.Visible = False
    Me.txtCarregando.Visible = True
    Me.Refresh
    PrintConnect
    Rpt.Connect = Conexao
    SQL = "Select * From tblMovimentoCaixa Order by Protocolo"
    Screen.MousePointer = 11
    Rpt.ReportFileName = App.Path & "\Mov_Cancelado.rpt"
    Rpt.Destination = crptToWindow
    Rpt.SQLQuery = SQL
    Rpt.Formulas(0) = "Data='" & Me.txtDataInicio & "'"
    Rpt.Formulas(1) = "Data_Fim='" & Me.txtDataFim & "'"
    Rpt.WindowState = crptMaximized
    Rpt.CopiesToPrinter = 1
    Me.txtCarregando.Visible = False
    Rpt.Action = 1
    Screen.MousePointer = 1
    DB.Execute "Delete From tblMovimentoCaixa"

Exit Sub

Erro:
    If TransacaoAtiva Then
        DB.RollbackTrans
    End If
    MsgBox "Erro de Sistema. " & Err.Description, vbCritical
    DB.Execute "Delete From tblMovimentoCaixa"
    Me.cmdProcessando.Visible = False
    Me.txtCarregando.Visible = False
End Sub


