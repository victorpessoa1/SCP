Private Sub cmdCustasProtesto_Click()
On Error GoTo Erro
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim RSMovCaixa As ADODB.Recordset
    Set RSMovCaixa = New ADODB.Recordset
    Dim TransacaoAtiva As Boolean
    TransacaoAtiva = False
    Dim RSAponta As ADODB.Recordset
    Set RSAponta = New ADODB.Recordset


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

    Me.cmdCustasProtesto.Enabled = False
    Me.cmdProcessando.Visible = True
    Me.Refresh

    RS.Open "SELECT t.*, f.* " & _
    "FROM tblFinanceiro AS f " & _
    "LEFT JOIN tblTitulo AS t ON t.Protocolo_Cartorio = f.Protocolo " & _
    "Where (t.Data_Pagamento BETWEEN '" & Format(Me.txtDataInicio, "yyyy/mm/dd") & "' AND '" & Format(Me.txtDataFim, "yyyy/mm/dd") & "' AND f.Tipo_Ocorrencia = '1' AND t.CustasProtesto = '1' AND t.Anulado = '0' AND t.Aguardando = '0' AND t.CancelaBanco = '0' AND f.Data_Pagamento BETWEEN '" & Format(Me.txtDataInicio, "yyyy/mm/dd") & "' AND '" & Format(Me.txtDataFim, "yyyy/mm/dd") & "' AND (f.Baixa_Lote = 0 OR f.Baixa_Lote IS NULL) AND f.custas <> 0) " & _
    " OR (f.Rec_Data BETWEEN '" & Format(Me.txtDataInicio, "yyyy/mm/dd") & "' AND '" & Format(Me.txtDataFim, "yyyy/mm/dd") & "' AND f.Tipo_Ocorrencia = '1' AND t.CustasProtesto = '1' AND t.CancelaBanco = '1' AND f.Rec_Canc = '1' AND f.Baixa_Lote = '1' AND f.custas <> 0) ", DB, adOpenDynamic

    If RS.RecordCount = 0 Then
        Me.cmdProcessando.Visible = False
        Me.Refresh
        MsgBox "Sem Dados para exibir!", vbInformation
        RS.Close
        Exit Sub
    End If
'        nResp = MsgBox("Confirma a impressão de " & RS.RecordCount & " Títulos Pagos?", vbQuestion + vbYesNo)
        nResp = vbYes
        If nResp = vbYes Then
            Do While Not RS.EOF
                    If IsNull(RS!Estorno) Or RS!Estorno = False Then
                        DB.BeginTrans

                        RSAponta.Open "Select * From tblApontamento", DB, adOpenDynamic
                        Calcula_Atos RSAponta
                        If IsNull(RS!FRJ) Then
                            FRJ = 0
                            FRC = 0
                            vDistribuidor = RS!Valor_Distrib
                            Custas = RS!Custas
                        Else
                            FRJ = RS!FRJ - frjDistrib
                            FRC = RS!FRC - frcDistrib
                            vDistribuidor = RSAponta!Distribuidor
                            Custas = RS!Custas
                        End If

                        If IsNull(RS!TaxaCartao) Then
                            TxCartao = 0
                        Else
                            TxCartao = RS!TaxaCartao
                        End If

                        If RS!CancelaBanco = True Then
                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Valor_Juros,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,FRJ,FRC,TxBanco,Texto,ISS, Cartao) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & 0 & "','" & 0 & "','" & 0 & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & nNosso_Num & "','" & Replace(Portador, "'", " ") & "','" & 0 & "','" & _
                            0 & "','" & 0 & "','" & 0 & "','" & RS!Usuario & "','" & Replace(0, ",", ".") & "','" & RS!Hora & "','" & 0 & "','" & 0 & "','" & 0 & "','" & "BANCO" & "','" & Replace(RS!ISS, ",", ".") & "','" & RS!Cartao & "')")
                        Else
                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Valor_Juros,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,FRJ,FRC,TxBanco,ISS, Cartao) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(Custas, ",", ".") & "','" & Replace(RS!Valor_Selo, ",", ".") & "','" & Replace(vDistrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & nNosso_Num & "','" & Replace(Portador, "'", " ") & "','" & Replace(RS!Valor_Juros, ",", ".") & "','" & _
                            Replace(RS!Valor_Mora, ",", ".") & "','" & Replace(RS!V_Multa, ",", ".") & "','" & Replace(RS!Valor_CPMF, ",", ".") & "','" & RS!Usuario & "','" & Replace(0, ",", ".") & "','" & RS!Hora & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & Replace(TxCartao, ",", ".") & "','" & Replace(RS!ISS, ",", ".") & "','" & RS!Cartao & "')")
                        End If

                        RSAponta.Close
                        DB.CommitTrans
                        RS.MoveNext
                    Else
                        RS.MoveNext
                    End If
                Loop
                RS.Close
        Else
            Me.cmdProcessando.Visible = False
            MsgBox "Impressão Cancelada!", vbInformation
            RS.Close
            DB.Execute "Delete From tblMovimentoCaixa"
            Exit Sub
        End If

'<<< IMPRESSÃO >>>
            Me.cmdProcessando.Visible = False
            Me.txtCarregando.Visible = True
            Me.Refresh
            PrintConnect
            Rpt.Connect = Conexao
            SQL = "Select * From tblMovimentoCaixa Order by Protocolo"
            Screen.MousePointer = 11
            Rpt.ReportFileName = App.Path & "\Mov_CustasProtesto.rpt"
            Rpt.Destination = crptToWindow
            Rpt.SQLQuery = SQL
            Rpt.Formulas(0) = "Data='" & Me.txtData & "'"
            Rpt.WindowState = crptMaximized
            Rpt.CopiesToPrinter = 1
            Me.txtCarregando.Visible = False
            Rpt.Action = 1
            Screen.MousePointer = 1
'<<< FIM IMPRESSÃO >>>
            DB.Execute "Delete From tblMovimentoCaixa"
Exit Sub
Erro:
    DB.RollbackTrans
    MsgBox "Erro de Sistema. " & Err.Description, vbCritical
    DB.Execute "Delete From tblMovimentoCaixa"
    Me.cmdProcessando.Visible = False
    Me.txtCarregando.Visible = False

End Sub




