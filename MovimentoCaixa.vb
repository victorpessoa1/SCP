Private Sub cmdCancelados_Click()
On Error GoTo Erro
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim RSFinan As ADODB.Recordset
    Set RSFinan = New ADODB.Recordset
    Dim RSMovCaixa As ADODB.Recordset
    Set RSMovCaixa = New ADODB.Recordset
    DB.Execute "Delete From tblMovimentoCaixa"
    RSMovCaixa.Open "Select * From tblMovimentoCaixa", DB, adOpenDynamic
        If RSMovCaixa.RecordCount > 0 Then
            nMC = MsgBox("O Relatório está em uso. Liberar?.", vbQuestion + vbYesNo)
            If nMC = vbYes Then
                DB.BeginTrans
                DB.Execute "Delete From tblMovimentoCaixa"
                DB.CommitTrans
            Else
                RSMovCaixa.Close
            End If
            Exit Sub
        End If
    RSMovCaixa.Close

    Me.cmdCancelados.Enabled = False
    Me.cmdProcessando.Visible = True
    Me.Refresh
    RS.Open "Select * From tblTitulo Where Data_Cancelado Between '" & Format(Me.txtDataInicio, "mm/dd/yyyy") & "' AND '" & Format(Me.txtDataFim, "mm/dd/yyyy") & "' AND Tipo_Ocorrencia='" & "A" & "' AND Anulado='" & 0 & "' AND Aguardando='" & 0 & "'", DB, adOpenDynamic
    If RS.RecordCount = 0 Then
        Me.cmdProcessando.Visible = False
        Me.Refresh
        MsgBox "Sem Dados para exibir!", vbInformation
        RS.Close
        Exit Sub
    Else
'        nResp = MsgBox("Confirma a impressão de " & RS.RecordCount & " Títulos Cancelados?", vbQuestion + vbYesNo)
        nResp = vbYes
        If nResp = vbYes Then
            Do While Not RS.EOF
'                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Data_Cancelado='" & Format(Me.txtData, "mm/dd/yyyy") & "' AND Baixa_Lote Is Null", DB, adOpenDynamic
                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Data_Cancelado Between '" & Format(Me.txtDataInicio, "mm/dd/yyyy") & "' AND '" & Format(Me.txtDataFim, "mm/dd/yyyy") & "'", DB, adOpenDynamic
                If RSFinan.RecordCount = 0 Then
'                    Me.cmdProcessando.Visible = False
'                    Me.Refresh
'                    MsgBox "O Protocolo " & RS!Protocolo_Cartorio & " não foi lançado no Financeiro.", vbInformation
'                    RSFinan.Close
'                    RS.Close
'                    Exit Sub
                End If

                Do While Not RSFinan.EOF
                    If IsNull(RSFinan!Estorno) Or RSFinan!Estorno = False Then
                        DB.BeginTrans
                        If IsNull(RSFinan!TaxaCartao) Then
                            TxCartao = 0
                        Else
                            TxCartao = RSFinan!TaxaCartao
                        End If

                        If RSFinan!Codigo = 1 Then
                            TxBanco = 1
                        Else
                            TxBanco = Replace(TxCartao, ",", ".")
                        End If

                        If IsNull(RSFinan!FRJ) Then
                            FRJ = o
                            FRC = 0
                        Else
                            FRJ = RSFinan!FRJ
                            FRC = RSFinan!FRC
                        End If
                        If IsNull(RSFinan!Cartao) Or RSFinan!Cartao = False Then
                            Cartao = 0
                        Else
                            Cartao = 1
                        End If

                        If RS!CancelaBanco = True Then
                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Livro,Folha,Usuario,Texto,TxBanco,ISS,Data_Protestado,FRJ,FRC,Cartao) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & 0 & "','" & 0 & "','" & 0 & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "','" & RS!Livro & "','" & RS!Pagina & "','" & RSFinan!Usuario & "','" & "BANCO" & "','" & TxBanco & "','" & 0 & "','" & Format(RS!data_protestado, "mm/dd/yyyy") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "'," & Cartao & ")")
                        Else
                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Livro,Folha,Usuario,TxBanco,ISS,Data_Protestado,FRJ,FRC,Cartao) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(RSFinan!Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "','" & RS!Livro & "','" & RS!Pagina & "','" & RSFinan!Usuario & "','" & TxBanco & "','" & Replace(RSFinan!ISS, ",", ".") & _
                            "','" & Format(RS!data_protestado, "mm/dd/yyyy") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "'," & Cartao & ")")
                        End If
                        
                        DB.CommitTrans
                        RSFinan.MoveNext
                    Else
                        RSFinan.MoveNext
                    End If
                Loop
                RSFinan.Close
                RS.MoveNext
            Loop
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
            
'<<< FIM IMPRESSÃO >>>
            DB.Execute "Delete From tblMovimentoCaixa"
    End If
Exit Sub
Erro:
    DB.RollbackTrans
    MsgBox "Erro de Sistema. " & Err.Description, vbCritical
    DB.Execute "Delete From tblMovimentoCaixa"
    Me.cmdProcessando.Visible = False
    Me.txtCarregando.Visible = False
    

End Sub