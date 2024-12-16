    Dim cAto As Integer
    Dim vAponta As Currency
    Dim vPago As Currency
    Dim vIntima As Currency
    Dim vCanc As Currency
    Dim vDistrib As Currency
    Dim vCpd As Currency
    Dim vProt As Currency
    Dim vRet As Currency
    Dim vCtProt As Currency
    Dim vEdital As Currency
    
    Dim eAponta As Currency
    Dim ePago As Currency
    Dim eIntima As Currency
    Dim eCanc As Currency
    Dim eDistrib As Currency
    Dim eCpd As Currency
    Dim eProt As Currency
    Dim eRet As Currency
    Dim eCtProt As Currency
    Dim eEdital As Currency
    
    Dim frcAponta As Currency
    Dim frcPago As Currency
    Dim frcIntima As Currency
    Dim frcCanc As Currency
    Dim frcDistrib As Currency
    Dim frcCpd As Currency
    Dim frcProt As Currency
    Dim frcRet As Currency
    Dim frcCtProt As Currency
    Dim frcEdital As Currency
    
    Dim frjAponta As Currency
    Dim frjPago As Currency
    Dim frjIntima As Currency
    Dim frjCanc As Currency
    Dim frjDistrib As Currency
    Dim frjCpd As Currency
    Dim frjProt As Currency
    Dim frjRet As Currency
    Dim frjCtProt As Currency
    Dim frjEdital As Currency

Private Sub Calendario_DateClick(ByVal DateClicked As Date)
If Me.txtDataInicio = "" Then
    Me.txtDataInicio = Format(Me.Calendario.Value, "dd/mm/yyyy")
    Me.txtData = Format(Me.Calendario.Value, "dd/mm/yyyy")
Else
    Me.txtDataFim = Format(Me.Calendario.Value, "dd/mm/yyyy")
    Me.cmdCancelados.Enabled = True
    Me.cmdCertidao.Enabled = True
    Me.cmdPagamentos.Enabled = True
    Me.cmdProtestados.Enabled = True
    Me.cmdRetirados.Enabled = True
    Me.cmdRetiradosPart.Enabled = True
    Me.cmdCustasProtesto.Enabled = True
    Me.cmdDepositoPrevio.Enabled = True
    Me.cmdCustas.Enabled = True
    Me.cmdFechamento.Enabled = True
    Me.cmdCDA.Enabled = True
    Me.cmdBoletos.Enabled = True
End If
End Sub

Private Sub cmdBoletos_Click()
On Error GoTo Erro
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim RSFinan As ADODB.Recordset
    Set RSFinan = New ADODB.Recordset
    Dim RSMovCaixa As ADODB.Recordset
    Set RSMovCaixa = New ADODB.Recordset
    Dim Aponta As ADODB.Recordset
    Set Aponta = New ADODB.Recordset
    Dim DataPag As Date
    
    Aponta.Open "Select * FROM tblApontamento", DB, adOpenDynamic
    
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
    
    Me.cmdBoletos.Enabled = False
    Me.cmdProcessando.Visible = True
    Me.Refresh
    
'<<< Pagamentos >>>
    RS.Open "Select * From tblTitulo Where Data_Pagamento='" & Format(Me.txtData, "mm/dd/yyyy") & "' AND CustasProtesto Is Null " & " AND Baixado='" & 1 & "' AND Anulado='" & 0 & "' AND CancelaBanco='" & 0 & "' AND Aguardando='" & 0 & "'", DB, adOpenDynamic
    If RS.RecordCount = 0 Then
        Me.cmdProcessando.Visible = False
        Me.Refresh
        MsgBox "Sem Dados para exibir!", vbInformation
        RS.Close
        Exit Sub
    Else
            Do While Not RS.EOF
                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND TED >'" & 0 & "' AND Data_Pagamento='" & Format(Me.txtData, "mm/dd/yyyy") & "'", DB, adOpenDynamic
                If RSFinan.RecordCount = 0 Then
                    Me.cmdProcessando.Visible = False
                    Me.Refresh
                    'MsgBox "O Protocolo " & RS!Protocolo_Cartorio & " não foi lançado no Financeiro.", vbInformation
'                    RSFinan.Close
'                    RS.MoveNext
'                    RS.Close
                    'Exit Sub
                Else
                                
                                
Calcula_Atos Aponta
                                
                Do While Not RSFinan.EOF
                    If IsNull(RSFinan!Estorno) Or RSFinan!Estorno = False Then
                        DB.BeginTrans
                        DataPag = "01/01/2020"
                        If IsNull(RSFinan!TED) Then
                            TED = 0
                        Else
                            TED = RSFinan!TED
                        End If
                        
                        If Format(RS!Data_Pagamento, "mm/dd/yyyy") >= DataPag Then
                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Valor_Juros,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,FRJ,FRC,TxBanco,TxCartao,Tipo_Baixa,ISS) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(vDistrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Replace(Portador, "'", " ") & "','" & Replace(RSFinan!Valor_Juros, ",", ".") & "','" & Replace(RSFinan!Valor_Mora, ",", ".") & "','" & _
                            Replace(RSFinan!V_Multa, ",", ".") & "','" & Replace(RSFinan!Valor_CPMF, ",", ".") & "','" & RSFinan!Usuario & "','" & Replace(0, ",", ".") & "','" & RSFinan!Hora & "','" & Replace(RSFinan!FRJ - frjDistrib, ",", ".") & "','" & Replace(RSFinan!FRC - frcDistrib, ",", ".") & "','" & Replace(TED, ",", ".") & "','" & 1 & "','" & "Pagamento" & "','" & Replace(RSFinan!ISS, ",", ".") & "')")
                        Else
                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Valor_Juros,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,FRJ,FRC,TxBanco,TxCartao,Tipo_Baixa,ISS) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(RSFinan!Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Replace(Portador, "'", " ") & "','" & Replace(RSFinan!Valor_Juros, ",", ".") & "','" & Replace(RSFinan!Valor_Mora, ",", ".") & "','" & _
                            Replace(RSFinan!V_Multa, ",", ".") & "','" & Replace(RSFinan!Valor_CPMF, ",", ".") & "','" & RSFinan!Usuario & "','" & Replace(0, ",", ".") & "','" & RSFinan!Hora & "','" & Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(TED, ",", ".") & "','" & 1 & "','" & "Pagamento" & "','" & Replace(RSFinan!ISS, ",", ".") & "')")
                        End If
                        DB.CommitTrans
                        RSFinan.MoveNext
                    Else
                        RSFinan.MoveNext
                    End If
                Loop
                End If
                RSFinan.Close
                RS.MoveNext
            Loop
            
'<<< Cancelamentos >>>
    RS.Close
    RS.Open "Select * From tblTitulo Where Data_Cancelado='" & Format(Me.txtData, "mm/dd/yyyy") & "' AND Tipo_Ocorrencia='" & "A" & "' AND Anulado='" & 0 & "' AND Aguardando='" & 0 & "'", DB, adOpenDynamic
    If RS.RecordCount = 0 Then
        Me.cmdProcessando.Visible = False
        Me.Refresh
        MsgBox "Sem Dados para exibir!", vbInformation
        RS.Close
        Exit Sub
    Else
            Do While Not RS.EOF
                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Data_Cancelado='" & Format(Me.txtData, "mm/dd/yyyy") & "' AND Codigo = '" & "1" & "'", DB, adOpenDynamic
                If RSFinan.RecordCount = 0 Then
                    Me.cmdProcessando.Visible = False
                    Me.Refresh
                Else
                                
                Do While Not RSFinan.EOF
                    If IsNull(RSFinan!Estorno) Or RSFinan!Estorno = False Then
                        DB.BeginTrans
                        If IsNull(RSFinan!TaxaCartao) Then
                            TxCartao = 0
                        Else
                            TxCartao = RSFinan!TaxaCartao
                        End If
                        
                            TxBanco = 1
                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Valor_Juros,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,FRJ,FRC,TxBanco,TxCartao,Tipo_Baixa) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Replace(Portador, "'", " ") & "','" & Replace(RSFinan!Valor_Juros, ",", ".") & "','" & Replace(RSFinan!Valor_Mora, ",", ".") & "','" & _
                            Replace(RSFinan!V_Multa, ",", ".") & "','" & Replace(RSFinan!Valor_CPMF, ",", ".") & "','" & RSFinan!Usuario & "','" & Replace(0, ",", ".") & "','" & RSFinan!Hora & "','" & Replace(RSFinan!FRJ, ",", ".") & "','" & Replace(RSFinan!FRC, ",", ".") & "','" & Replace(TED, ",", ".") & "','" & 1 & "','" & "Cancelamento" & "')")
                        
                        DB.CommitTrans
                        RSFinan.MoveNext
                    Else
                        RSFinan.MoveNext
                    End If
                Loop
                End If
                RSFinan.Close
                RS.MoveNext
            Loop
        End If

'<<< Custas de Protesto >>>
    RS.Close
    RS.Open "Select * From tblTitulo Where Data_Cancelado='" & Format(Me.txtData, "mm/dd/yyyy") & "' AND Tipo_Ocorrencia='" & "A" & "' AND Anulado='" & 0 & "' AND Aguardando='" & 0 & "' AND Especie_Tit='" & "CDA" & "'", DB, adOpenDynamic
    If RS.RecordCount = 0 Then
        Me.cmdProcessando.Visible = False
        Me.Refresh
    Else
''        nResp = MsgBox("Confirma a impressão de " & RS.RecordCount & " CDA's?", vbQuestion + vbYesNo)
'        nResp = vbYes
'        If nResp = vbYes Then
            Do While Not RS.EOF
                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Especie_Tit='" & "CDA" & "' AND Codigo='" & "1" & "'", DB, adOpenDynamic
                If RSFinan.RecordCount = 0 Then
                    Me.cmdProcessando.Visible = False
                    Me.Refresh
'                    MsgBox "O Protocolo " & RS!Protocolo_Cartorio & " não foi lançado no Financeiro.", vbInformation
'                    RSFinan.Close
'                    RS.Close
'                    Exit Sub
                End If
                                                
                Do While Not RSFinan.EOF
                If RSFinan!Tipo_Ocorrencia = "1" Then
                    If IsNull(RSFinan!Estorno) Or RSFinan!Estorno = False Then
                        DB.BeginTrans
                        If RSFinan!Codigo = 1 Then
                            TxBanco = "0,00"
                        Else
                            TxBanco = 0
                        End If
                        
                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Valor_Juros,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,FRJ,FRC,TxBanco,TxCartao,Tipo_Baixa) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(RSFinan!Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Replace(Portador, "'", " ") & "','" & Replace(RSFinan!Valor_Juros, ",", ".") & "','" & Replace(RSFinan!Valor_Mora, ",", ".") & "','" & _
                            Replace(RSFinan!V_Multa, ",", ".") & "','" & Replace(RSFinan!Valor_CPMF, ",", ".") & "','" & RSFinan!Usuario & "','" & Replace(0, ",", ".") & "','" & RSFinan!Hora & "','" & Replace(RSFinan!FRJ - frjDistrib, ",", ".") & "','" & Replace(RSFinan!FRC - frcDistrib, ",", ".") & "','" & Replace(TxBanco, ",", ".") & "','" & 1 & "','" & "Cancelamento" & "')")
                        
'                        If RS!CancelaBanco = True Then
'                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Livro,Folha,Usuario,Texto,TxBanco) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & 0 & "','" & 0 & "','" & 0 & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "','" & RS!Livro & "','" & RS!Pagina & "','" & RSFinan!Usuario & "','" & "BANCO" & "','" & TxBanco & "')")
'                        Else
'                            If RSFinan!Pagar = 0 Then
'                                DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Livro,Folha,Usuario,TxBanco) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(RSFinan!Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "','" & RS!Livro & "','" & RS!Pagina & "','" & RSFinan!Usuario & "','" & TxBanco & "')")
'                            Else
'                                DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Livro,Folha,Usuario,TxBanco) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Pagar - RSFinan!Valor_Selo - RSFinan!Valor_Distrib, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(RSFinan!Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "','" & RS!Livro & "','" & RS!Pagina & "','" & RSFinan!Usuario & "','" & TxBanco & "')")
'                            End If
'                        End If
                        
                        DB.CommitTrans
                        RSFinan.MoveNext
                    Else
                        RSFinan.MoveNext
                    End If
                Else
                    RSFinan.MoveNext
                End If
                
                Loop
                RSFinan.Close
                RS.MoveNext
            Loop
'        Else
'            Me.cmdProcessando.Visible = False
'            MsgBox "Impressão Cancelada!", vbInformation
'            RS.Close
'            DB.Execute "Delete From tblMovimentoCaixa"
'            Exit Sub
        End If



'<<< IMPRESSÃO >>>
            Me.cmdProcessando.Visible = False
            Me.txtCarregando.Visible = True
            Me.Refresh
            PrintConnect
            Rpt.Connect = Conexao
            SQL = "Select * From tblMovimentoCaixa WHERE TxCartao='" & 1 & "' Order by Protocolo"
            Screen.MousePointer = 11
            Rpt.ReportFileName = App.Path & "\Mov_Boletos.rpt"
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
            Aponta.Close
            RS.Close
    End If
Exit Sub
Erro:
'    DB.RollbackTrans
    MsgBox "Erro de Sistema. " & Err.Description, vbCritical
    DB.Execute "Delete From tblMovimentoCaixa"
    Me.cmdProcessando.Visible = False
    Me.txtCarregando.Visible = False
    Aponta.Close

End Sub

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

Private Sub cmdCDA_Click()
On Error GoTo Erro
    Dim vDistrib As Currency
    Dim eDistrib As Currency
    Dim frcDistrib As Currency
    Dim frjDistrib As Currency
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim RSFinan As ADODB.Recordset
    Set RSFinan = New ADODB.Recordset
    Dim RSMovCaixa As ADODB.Recordset
    Set RSMovCaixa = New ADODB.Recordset
    Dim RSAponta As ADODB.Recordset
    Set RSAponta = New ADODB.Recordset
    
    RSAponta.Open "Select * From tblApontamento", DB, adOpenDynamic
    vDistrib = RSAponta!Distribuidor
    frjDistrib = Format(vDistrib * 0.15, "0.00")
    frcDistrib = Format(vDistrib * 0.025, "0.00")
    txDistrib = frjDistrib + frcDistrib
    
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
    
    Me.cmdCDA.Enabled = False
    Me.cmdProcessando.Visible = True
    Me.Refresh
    RS.Open "Select * From tblTitulo Where Data_Cancelado='" & Format(Me.txtData, "mm/dd/yyyy") & "' AND Tipo_Ocorrencia='" & "A" & "' AND Anulado='" & 0 & "' AND Aguardando='" & 0 & "' AND Especie_Tit='" & "CDA" & "'", DB, adOpenDynamic
    If RS.RecordCount = 0 Then
        Me.cmdProcessando.Visible = False
        Me.Refresh
        MsgBox "Sem Dados para exibir!", vbInformation
        RS.Close
        Exit Sub
    Else
'        nResp = MsgBox("Confirma a impressão de " & RS.RecordCount & " CDA's?", vbQuestion + vbYesNo)
        nResp = vbYes
        If nResp = vbYes Then
            Do While Not RS.EOF
                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Especie_Tit='" & "CDA" & "'", DB, adOpenDynamic
                If RSFinan.RecordCount = 0 Then
                    Me.cmdProcessando.Visible = False
                    Me.Refresh
                    MsgBox "O Protocolo " & RS!Protocolo_Cartorio & " não foi lançado no Financeiro.", vbInformation
                    RSFinan.Close
                    RS.Close
                    Exit Sub
                End If
                                                
                Do While Not RSFinan.EOF
                If RSFinan!Tipo_Ocorrencia = "1" Then
                    If IsNull(RSFinan!Estorno) Or RSFinan!Estorno = False Then
                        DB.BeginTrans
                        If RSFinan!Codigo = 1 Then
                            TxBanco = 1
                        Else
                            TxBanco = 0
                        End If
                        
                        If IsNull(RSFinan!FRJ) Then
                            FRJ = 0
                            FRC = 0
                            vDistribuidor = RSFinan!Valor_Distrib
                            Custas = RSFinan!Custas
                        Else
                            FRJ = RSFinan!FRJ - frjDistrib
                            FRC = RSFinan!FRC - frcDistrib
                            vDistribuidor = RSAponta!Distribuidor
                            Custas = RSFinan!Custas
                        End If
                        
                        If IsNull(RSFinan!ISS) Then
                            ISS = 0
                        Else
                            ISS = RSFinan!ISS
                        End If
                        
                        If RS!CancelaBanco = True Then
                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Livro,Folha,Usuario,Texto,TxBanco,ISS,FRJ,FRC) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & 0 & "','" & 0 & "','" & 0 & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "','" & RS!Livro & "','" & RS!Pagina & "','" & RSFinan!Usuario & "','" & "BANCO" & "','" & TxBanco & "','" & _
                            Replace(ISS, ",", ".") & "','" & 0 & "','" & 0 & "')")
                        Else
                            If RSFinan!Pagar = 0 Then
                                DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Livro,Folha,Usuario,TxBanco,ISS,FRJ,FRC) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(vDistribuidor, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "','" & RS!Livro & "','" & RS!Pagina & "','" & RSFinan!Usuario & "','" & TxBanco & "','" & _
                                Replace(ISS, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "')")
                            Else
                                DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Livro,Folha,Usuario,TxBanco,ISS,FRJ,FRC) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Pagar - RSFinan!Valor_Selo - RSFinan!Valor_Distrib, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(RSFinan!Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "','" & RS!Livro & "','" & RS!Pagina & "','" & RSFinan!Usuario & "','" & TxBanco & "','" & _
                                Replace(ISS, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "')")
                            End If
                        End If
                        
                        DB.CommitTrans
                        RSFinan.MoveNext
                    Else
                        RSFinan.MoveNext
                    End If
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
            Rpt.ReportFileName = App.Path & "\Mov_CDA.rpt"
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
    End If
Exit Sub
Erro:
    DB.RollbackTrans
    MsgBox "Erro de Sistema. " & Err.Description, vbCritical
    DB.Execute "Delete From tblMovimentoCaixa"
    Me.cmdProcessando.Visible = False
    Me.txtCarregando.Visible = False
End Sub

Private Sub cmdCertidao_Click()
On Error GoTo Erro
    Dim RSFinan As ADODB.Recordset
    Set RSFinan = New ADODB.Recordset
    Dim RSMovCaixa As ADODB.Recordset
    Set RSMovCaixa = New ADODB.Recordset
    Dim RSAponta As ADODB.Recordset
    Set RSAponta = New ADODB.Recordset

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
    
    Me.cmdCertidao.Enabled = False
    Me.cmdProcessando.Visible = True
    Me.Refresh
    RSFinan.Open "SELECT * FROM tblFinanceiro INNER JOIN tblReqCertidao ON tblFinanceiro.Codigo = tblReqCertidao.Codigo WHERE Data_Certidao Between '" & Format(Me.txtDataInicio, "yyyy/mm/dd") & "' AND '" & Format(Me.txtDataFim, "yyyy/mm/dd") & "' AND tblReqCertidao.Pago = 1" & "AND VlrCertidao != 0", DB, adOpenDynamic
    If RSFinan.RecordCount = 0 Then
        Me.cmdProcessando.Visible = False
        Me.Refresh
        MsgBox "Sem Dados para exibir!", vbInformation
        RSFinan.Close
        Exit Sub
    Else
'        nResp = MsgBox("Confirma a impressão de " & RSFinan.RecordCount & " Certidões?", vbQuestion + vbYesNo)
        nResp = vbYes
        If nResp = vbYes Then
        RSAponta.Open "SELECT * FROM tblApontamento", DB, adOpenDynamic
            Do While Not RSFinan.EOF
                'If IsNull(RSFinan!Estorno) Or RSFinan!Estorno = False Then
                    If IsNull(RSFinan!VlrCertidao) Then
                        vCertidao = 0
                    Else
                        vCertidao = RSFinan!VlrCertidao
                    End If
                    If IsNull(RSFinan!Valor_Selo) Then
                        Valor_Selo = 0
                    Else
                        Valor_Selo = RSFinan!Valor_Selo
                    End If
                    If IsNull(RSFinan!TaxaCartao) Then
                        TxBanco = 0
                    Else
                        TxBanco = RSFinan!TaxaCartao
                    End If
                    If IsNull(RSFinan!Cartao) Or RSFinan!Cartao = False Then
                        Cartao = 0
                    Else
                        Cartao = 1
                    End If
                    If IsNull(RSFinan!FRJ) Then
                        FRJ = 0
                        FRC = 0
'                        ISS = 0
                    Else
                        FRJ = RSFinan!FRJ
                        FRC = RSFinan!FRC
                        ISS = RSFinan!ISS
                    End If
                    If IsNull(RSFinan!ISS) Then
                        ISS = 0
                    Else
                        ISS = RSFinan!ISS
                    End If
                    If RSFinan!Estorno = True Then
                        FRC = 0
                        FRJ = 0
                        ISS = 0
                    End If
                    If RSFinan!TipoCertidao = "NEGATIVA" Then
                    Else
                        DB.BeginTrans
                        DB.Execute "INSERT INTO tblMovimentoCaixa (Devedor, Num_Doc, Saldo, Selo, Usuario, Tipo_Certidao, Hora, TxBanco, FRJ, FRC, ISS, Cartao) " & _
                        "VALUES ('" & Replace(RSFinan!NomeCertidao, "'", "''") & "','" & RSFinan!DocCertidao & "','" & Replace(vCertidao, ",", ".") & "','" & Replace(Valor_Selo, ",", ".") & "','" & RSFinan!Usuario & "','" & RSFinan!TipoCertidao & "','" & RSFinan!Hora & "','" & Replace(TxBanco, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & Replace(ISS, ",", ".") & "','" & Cartao & "')"
                        DB.CommitTrans
                    End If
                    RSFinan.MoveNext
                'Else
                    'RSFinan.MoveNext
                'End If
            Loop
            RSFinan.Close
        Else
            Me.cmdProcessando.Visible = False
            MsgBox "Impressão Cancelada!", vbInformation
            RSFinan.Close
            DB.Execute "Delete From tblMovimentoCaixa"
            Exit Sub
        End If

'<<< IMPRESSÃO >>>
            Me.cmdProcessando.Visible = False
            Me.txtCarregando.Visible = True
            Me.Refresh
            PrintConnect
            Rpt.Connect = Conexao
            SQL = "Select * From tblMovimentoCaixa Order by Devedor"
            Screen.MousePointer = 11
            Rpt.ReportFileName = App.Path & "\Mov_Certidao.rpt"
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
            RSAponta.Close
    End If
Exit Sub
Erro:
    MsgBox "Erro de Sistema. " & Err.Description, vbCritical
    Me.cmdProcessando.Visible = False
    Me.txtCarregando.Visible = False
    DB.RollbackTrans
    DB.Execute "Delete From tblMovimentoCaixa"
End Sub

Private Sub cmdCustas_Click()
On Error GoTo Erro
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    Me.cmdCustas.Enabled = False
    Me.cmdProcessando.Visible = True
    Me.Refresh
    SQL = "Select * From tblDepositoPrevio Where Data='" & Format(Me.txtData, "mm/dd/yyyy") & "' AND Restituicao='" & 1 & "'"
    RS.Open SQL, DB, adOpenDynamic
    If RS.RecordCount = 0 Then
        Me.cmdProcessando.Visible = False
        Me.Refresh
        MsgBox "Sem Dados para exibir!", vbInformation
        RS.Close
        Exit Sub
    Else
'        If RS.RecordCount > 1 Then
'            nResp = MsgBox("Confirma a impressão de " & RS.RecordCount & " Custas de Distribuição?", vbQuestion + vbYesNo)
'        Else
'            nResp = MsgBox("Confirma a impressão de " & RS.RecordCount & " Custas de Distribuição?", vbQuestion + vbYesNo)
'        End If
        nResp = vbYes
        If nResp = vbYes Then
'<<< IMPRESSÃO >>>
            Me.cmdProcessando.Visible = False
            Me.txtCarregando.Visible = True
            Me.Refresh
            PrintConnect
            Rpt.Connect = Conexao
            Screen.MousePointer = 11
            Rpt.ReportFileName = App.Path & "\Mov_CustasDistribuidor.rpt"
            Rpt.Destination = crptToWindow
            Rpt.SQLQuery = SQL
            Rpt.Formulas(0) = "Data='" & Me.txtData & "'"
            Rpt.WindowState = crptMaximized
            Rpt.CopiesToPrinter = 1
            Me.txtCarregando.Visible = False
            Screen.MousePointer = 1
            Rpt.Action = 1
'<<< FIM IMPRESSÃO >>>
        Else
            Me.cmdProcessando.Visible = False
            MsgBox "Impressão Cancelada!", vbInformation
            RS.Close
            Exit Sub
        End If

    End If
Exit Sub
Erro:
    DB.RollbackTrans
    MsgBox "Erro de Sistema. " & Err.Description, vbCritical
    DB.Execute "Delete From tblMovimentoCaixa"
    Me.cmdProcessando.Visible = False
    Me.txtCarregando.Visible = False

End Sub

Private Sub cmdCustasProtesto_Click()
On Error GoTo Erro
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim RSFinan As ADODB.Recordset
    Set RSFinan = New ADODB.Recordset
    Dim RSMovCaixa As ADODB.Recordset
    Set RSMovCaixa = New ADODB.Recordset
    Dim RSAponta As ADODB.Recordset
    Set RSAponta = New ADODB.Recordset
    
    RSMovCaixa.Open "Select * From tblMovimentoCaixa", DB, adOpenDynamic
        If RSMovCaixa.RecordCount > 0 Then
            nMC = MsgBox("O Relatório está em uso. Liberar?", vbQuestion + vbYesNo)
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
    
    Me.cmdCustasProtesto.Enabled = False
    Me.cmdProcessando.Visible = True
    Me.Refresh
    RS.Open "Select * From tblTitulo Where Data_Pagamento='" & Format(Me.txtData, "mm/dd/yyyy") & "' AND CustasProtesto='" & 1 & "'AND Especie_Tit<>'" & "CDA" & "'AND CancelaBanco='" & 0 & "'", DB, adOpenDynamic
    If RS.RecordCount = 0 Then
        Me.cmdProcessando.Visible = False
        Me.Refresh
        MsgBox "Sem Dados para exibir!", vbInformation
        RS.Close
        Exit Sub
    Else
'        nResp = MsgBox("Confirma a impressão de " & RS.RecordCount & " Títulos Pagos?", vbQuestion + vbYesNo)
        nResp = vbYes
        If nResp = vbYes Then
            Do While Not RS.EOF
                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Data_Pagamento='" & Format(Me.txtData, "mm/dd/yyyy") & "'AND Baixa_Lote Is Null", DB, adOpenDynamic
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
                        
                        RSAponta.Open "Select * From tblApontamento", DB, adOpenDynamic
                        Calcula_Atos RSAponta
                        If IsNull(RSFinan!FRJ) Then
                            FRJ = 0
                            FRC = 0
                            vDistribuidor = RSFinan!Valor_Distrib
                            Custas = RSFinan!Custas
                        Else
                            FRJ = RSFinan!FRJ - frjDistrib
                            FRC = RSFinan!FRC - frcDistrib
                            vDistribuidor = RSAponta!Distribuidor
                            Custas = RSFinan!Custas
                        End If
                        
                        If IsNull(RSFinan!TaxaCartao) Then
                            TxCartao = 0
                        Else
                            TxCartao = RSFinan!TaxaCartao
                        End If
                        
                        If RS!CancelaBanco = True Then
                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Valor_Juros,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,FRJ,FRC,TxBanco,Texto,ISS) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & 0 & "','" & 0 & "','" & 0 & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & nNosso_Num & "','" & Replace(Portador, "'", " ") & "','" & 0 & "','" & _
                            0 & "','" & 0 & "','" & 0 & "','" & RSFinan!Usuario & "','" & Replace(0, ",", ".") & "','" & RSFinan!Hora & "','" & 0 & "','" & 0 & "','" & 0 & "','" & "BANCO" & "','" & Replace(RSFinan!ISS, ",", ".") & "')")
                        Else
                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Valor_Juros,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,FRJ,FRC,TxBanco,ISS) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(vDistrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & nNosso_Num & "','" & Replace(Portador, "'", " ") & "','" & Replace(RSFinan!Valor_Juros, ",", ".") & "','" & _
                            Replace(RSFinan!Valor_Mora, ",", ".") & "','" & Replace(RSFinan!V_Multa, ",", ".") & "','" & Replace(RSFinan!Valor_CPMF, ",", ".") & "','" & RSFinan!Usuario & "','" & Replace(0, ",", ".") & "','" & RSFinan!Hora & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & Replace(TxCartao, ",", ".") & "','" & Replace(RSFinan!ISS, ",", ".") & "')")
                        End If
                        
                        RSAponta.Close
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
    End If
Exit Sub
Erro:
    DB.RollbackTrans
    MsgBox "Erro de Sistema. " & Err.Description, vbCritical
    DB.Execute "Delete From tblMovimentoCaixa"
    Me.cmdProcessando.Visible = False
    Me.txtCarregando.Visible = False

End Sub

Private Sub cmdDepositoPrevio_Click()
On Error GoTo Erro
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim RSDep As ADODB.Recordset
    Set RSDep = New ADODB.Recordset
    
    Me.cmdDepositoPrevio.Enabled = False
    Me.cmdProcessando.Visible = True
    Me.Refresh
    SQL = "Select * From tblDepositoPrevio Where Data='" & Format(Me.txtData, "mm/dd/yyyy") & "' AND Restituicao='" & 0 & "'"
    RS.Open SQL, DB, adOpenDynamic
    If RS.RecordCount = 0 Then
        Me.cmdProcessando.Visible = False
        Me.Refresh
        MsgBox "Sem Dados para exibir!", vbInformation
        RS.Close
        Exit Sub
    Else
        If RS.RecordCount > 1 Then
            nResp = MsgBox("Confirma a impressão de " & RS.RecordCount & " Depósitos Prévios?", vbQuestion + vbYesNo)
        Else
            nResp = MsgBox("Confirma a impressão de " & RS.RecordCount & " Depósito Prévio?", vbQuestion + vbYesNo)
        End If
        
        If nResp = vbYes Then
        
        Do While Not RS.EOF
            RSDep.Open "SELECT * FROM tblTitulo WHERE Protocolo_Cartorio='" & RS!Protocolo & "'", DB, adOpenDynamic
            If IsNull(RSDep!data_protestado) And Not IsNull(RSDep!Data_Pagamento) Then
                Data_Baixa = RSDep!Data_Pagamento
                DB.Execute "UPDATE tblDepositoPrevio SET Data_Baixa='" & Format(Data_Baixa, "mm/dd/yyyy") & "' WHERE id='" & RS!ID & "'"
            End If
            If Not IsNull(RSDep!data_protestado) And IsNull(RSDep!Data_Pagamento) Then
                Data_Prot = RSDep!data_protestado
                DB.Execute "UPDATE tblDepositoPrevio SET Data_Protestado='" & Format(Data_Prot, "mm/dd/yyyy") & "' WHERE id='" & RS!ID & "'"
            End If
            If Not IsNull(RSDep!data_protestado) And Not IsNull(RSDep!Data_Pagamento) Then
                Data_Prot = RSDep!data_protestado
                Data_Baixa = RSDep!Data_Pagamento
                DB.Execute "UPDATE tblDepositoPrevio SET Data_Protestado='" & Format(Data_Prot, "mm/dd/yyyy") & "',Data_Baixa='" & Format(Data_Baixa, "mm/dd/yyyy") & "' WHERE id='" & RS!ID & "'"
            End If
                
            RS.MoveNext
            RSDep.Close
        Loop
        
'<<< IMPRESSÃO >>>
            Me.cmdProcessando.Visible = False
            Me.txtCarregando.Visible = True
            Me.Refresh
            PrintConnect
            Rpt.Connect = Conexao
            Screen.MousePointer = 11
            Rpt.ReportFileName = App.Path & "\Mov_DepositoPrevio.rpt"
            Rpt.Destination = crptToWindow
            Rpt.SQLQuery = SQL
            Rpt.Formulas(0) = "Data='" & Me.txtData & "'"
            Rpt.WindowState = crptMaximized
            Rpt.CopiesToPrinter = 1
            Me.txtCarregando.Visible = False
            Screen.MousePointer = 1
            Rpt.Action = 1
'<<< FIM IMPRESSÃO >>>
        Else
            Me.cmdProcessando.Visible = False
            MsgBox "Impressão Cancelada!", vbInformation
            RS.Close
            Exit Sub
        End If

    End If
Exit Sub
Erro:
    DB.RollbackTrans
    MsgBox "Erro de Sistema. " & Err.Description, vbCritical
    DB.Execute "Delete From tblMovimentoCaixa"
    Me.cmdProcessando.Visible = False
    Me.txtCarregando.Visible = False

End Sub

Private Sub cmdFechamento_Click()
Dim CrysApp As New CRAXDDRT.Application
Dim CrysRep As New CRAXDDRT.Report


Dim CRExportOptions As Object

CrysApp.LogOnServer "crdb_odbc.dll", "SISPROT", "SISPROT", "sa", "kartorio@2012"
Set CrysRep = CrysApp.OpenReport(App.Path & "\Fechamento.rpt")
    With CrysRep
        .EnableParameterPrompting = False
        .DiscardSavedData
        .ParameterFields(1).AddCurrentValue Format(Me.txtDataFim, "yyyy/mm/dd")
        .ParameterFields(2).AddCurrentValue Format(Me.txtDataInicio, "yyyy/mm/dd")
        .ReadRecords
    End With

Set CRExportOptions = CrysRep.ExportOptions
CRExportOptions.FormatType = crEFTPortableDocFormat
CRExportOptions.DestinationType = crEDTDiskFile
CrysRep.DisplayProgressDialog = False
'CrysRep.Export False
frmPrtCaixa.CRViewer1.ReportSource = CrysRep
frmPrtCaixa.CRViewer1.ViewReport
frmPrtCaixa.Show 1
Set CRExportOptions = Nothing

'On Error GoTo Erro
'    Dim RS As ADODB.Recordset
'    Set RS = New ADODB.Recordset
'    Dim RSFinan As ADODB.Recordset
'    Set RSFinan = New ADODB.Recordset
'    Dim RSMovCaixa As ADODB.Recordset
'    Set RSMovCaixa = New ADODB.Recordset
'    Dim Aponta As ADODB.Recordset
'    Set Aponta = New ADODB.Recordset
'    Dim Cartao As Boolean
'
'    DB.Execute "Delete From tblMovimentoCaixa"
'
'    RSMovCaixa.Open "Select * From tblMovimentoCaixa", DB, adOpenDynamic
'        If RSMovCaixa.RecordCount > 0 Then
'            MsgBox "O Relatório está em uso. Tente mais tarde.", vbInformation
'            RSMovCaixa.Close
'            Exit Sub
'        End If
'    RSMovCaixa.Close
'
'    Me.cmdFechamento.Enabled = False
'    Me.cmdProcessando.Visible = True
'    Me.Refresh
'
'    Aponta.Open "SELECT * FROM tblApontamento", DB, adOpenDynamic
'    Calcula_Atos Aponta
'
'    '<<< PAGAMENTO >>>
'    RS.Open "Select * From tblTitulo Where Data_Pagamento BETWEEN'" & Format(Me.txtDataInicio, "mm/dd/yyyy") & "'AND'" & Format(Me.txtDataFim, "mm/dd/yyyy") & "' AND CustasProtesto Is Null AND Baixado='" & 1 & "' AND Anulado='" & 0 & "' AND Aguardando='" & 0 & "' AND CancelaBanco='" & 0 & "'", DB, adOpenDynamic
'    If RS.RecordCount = 0 Then
'        Me.cmdProcessando.Visible = False
'        Me.Refresh
'        RS.Close
'    Else
'            Do While Not RS.EOF
''                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Data_Pagamento='" & Format(Me.txtData, "mm/dd/yyyy") & "'AND Baixa_Lote Is Null", DB, adOpenDynamic
'                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Data_Pagamento BETWEEN'" & Format(Me.txtDataInicio, "mm/dd/yyyy") & "'AND'" & Format(Me.txtDataFim, "mm/dd/yyyy") & "'", DB, adOpenDynamic
'                If RSFinan.RecordCount = 0 Then
''                    Me.cmdProcessando.Visible = False
''                    Me.Refresh
''                    MsgBox "O Protocolo " & RS!Protocolo_Cartorio & " não foi lançado no Financeiro.", vbInformation
''                    RSFinan.Close
''                    RS.Close
''                    Exit Sub
'                End If
'
'                Do While Not RSFinan.EOF
'                    If IsNull(RSFinan!Estorno) Or RSFinan!Estorno = False Then
'                        DB.BeginTrans
'
'                        If IsNull(RSFinan!TED) Then
'                            TED = 0
'                        Else
'                            TED = RSFinan!TED
'                        End If
'
'                        If RSFinan!Codigo = "1" Then
'                            TxBanco = 1
'                        Else
'                            TxBanco = 0
'                        End If
'
'                        If IsNull(RSFinan!TaxaCartao) Then
'                            TxBanco = 0
'                        Else
'                            TxBanco = RSFinan!TaxaCartao
'                        End If
'                        If IsNull(RSFinan!FRJ) Then
'                            FRJ = 0
'                            FRC = 0
'                        Else
'                            FRJ = RSFinan!FRJ
'                            FRC = RSFinan!FRC
'                        End If
'
'                        If IsNull(RSFinan!Cartao) Or RSFinan!Cartao = False Then
'                            Cartao = 0
'                        Else
'                            Cartao = 1
'                        End If
'
'                        Valor_Distrib = RSFinan!Valor_Distrib - frjDistrib - frcDistrib
'                        Custas = RSFinan!Custas + frjDistrib + frcDistrib
'
'                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Valor_Juros,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,Tipo_Baixa,TxBanco,ISS,FRJ,FRC,Cartao) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & nNosso_Num & "','" & Replace(Portador, "'", " ") & "','" & Replace(RSFinan!Valor_Juros, ",", ".") & "','" & Replace(RSFinan!Valor_Mora, ",", ".") & "','" & Replace(RSFinan!V_Multa, ",", ".") & "','" & _
'                            Replace(RSFinan!Valor_CPMF, ",", ".") & "','" & RSFinan!Usuario & "','" & Replace(0, ",", ".") & "','" & RSFinan!Hora & "','" & "Custas Pagto" & "','" & Replace(TED, ",", ".") & "','" & Replace(RSFinan!ISS, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & Cartao & "')")
'
'                        DB.CommitTrans
'                        RSFinan.MoveNext
'                    Else
'                        RSFinan.MoveNext
'                    End If
'                Loop
'                RSFinan.Close
'                RS.MoveNext
'            Loop
'            RS.Close
'    End If
'
'    '<<< CANCELADOS >>>
'
'    RS.Open "SELECT tblTitulo.Especie_Tit,Protocolo_Cartorio, Num_Titulo, Vencimento, Data_Apresenta, Origem, Saldo, Custas, tblFinanceiro.Valor_Selo as Valor_Selo, Estorno, TaxaCartao, Codigo, FRJ, FRC, ISS,Cartao, tblTitulo.CancelaBanco as CancelaBanco, tblTitulo.Devedor as Devedor, Sacador, tblTitulo.Valor_Juros as Valor_Juros, Valor_Mora, tblFinanceiro.V_Multa as V_Multa, tblFinanceiro.Valor_CPMF as Valor_CPMF, tblFinanceiro.Valor_Distrib as Valor_Distrib, Hora, Usuario FROM tblTitulo INNER JOIN tblFinanceiro ON tblFinanceiro.Protocolo = tblTitulo.Protocolo_Cartorio WHERE tblTitulo.Data_Cancelado BETWEEN '" & Format(Me.txtDataInicio.Value, "mm/dd/yyyy") & "' AND '" & Format(Me.txtDataFim.Value, "mm/dd/yyyy") & "' AND tblTitulo.Tipo_Ocorrencia = 'A' AND Anulado = '0' AND tblTitulo.CancelaBanco = '0' AND tblFinanceiro.Data_Cancelado BETWEEN '" _
'    & Format(Me.txtDataInicio.Value, "mm/dd/yyyy") & "' AND '" & Format(Me.txtDataFim.Value, "mm/dd/yyyy") & "' AND Custas != 0", DB, adOpenDynamic
'
'    If RS.RecordCount = 0 Then
'        Me.cmdProcessando.Visible = False
'        Me.Refresh
'        RS.Close
'    Else
'            Do While Not RS.EOF
'
'                    If IsNull(RS!Estorno) Or RS!Estorno = False Then
'                        DB.BeginTrans
'                        If IsNull(RS!TaxaCartao) Then
'                            TxCartao = 0
'                        Else
'                            TxCartao = RS!TaxaCartao
'                        End If
'
'                        If RS!Codigo = "1" Then
'                            TxBanco = 1
'                        Else
'                            TxBanco = Replace(TxCartao, ",", ".")
'                        End If
'
'                        If IsNull(RS!FRJ) Then
'                            FRJ = 0
'                            FRC = 0
'                        Else
'                            FRJ = RS!FRJ
'                            FRC = RS!FRC
'                        End If
'
'                        If IsNull(RS!Cartao) Or RS!Cartao = False Then
'                            Cartao = 0
'                        Else
'                            Cartao = 1
'                        End If
'
'                        If IsNull(RS!Origem) Then
'                            Origem = 0
'                        Else
'                            Origem = RS!Origem
'                        End If
'
'                        If RS!CancelaBanco = True Then
'                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Valor_Juros,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,Texto,Tipo_Baixa,TxBanco,ISS,FRJ,FRC,Cartao) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & RS!Especie_Tit & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RS!Custas, ",", ".") & "','" & Replace(RS!Valor_Selo, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & nNosso_Num & "','" & Replace(RS!Valor_Juros, ",", ".") & _
'                            "','" & Replace(RS!Valor_Mora, ",", ".") & "','" & Replace(RS!V_Multa, ",", ".") & "','" & Replace(RS!Valor_CPMF, ",", ".") & "','" & RS!Usuario & "','" & Replace(RS!Valor_Distrib, ",", ".") & "','" & RS!Hora & "','" & "BANCO" & "','" & "Cancelados" & "','" & TxBanco & "','" & Replace(RS!ISS, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & Cartao & "')")
'                        Else
'                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,Tipo_Baixa,TxBanco,ISS,FRJ,FRC,Cartao) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & RS!Especie_Tit & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RS!Custas + TxCartao, ",", ".") & "','" & Replace(RS!Valor_Selo, ",", ".") & "','" & Replace(RS!Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & Replace(RS!Valor_Mora, ",", ".") & _
'                            "','" & Replace(RS!V_Multa, ",", ".") & "','" & Replace(RS!Valor_CPMF, ",", ".") & "','" & RS!Usuario & "','" & Replace(0, ",", ".") & "','" & RS!Hora & "','" & "Cancelados" & "','" & TxBanco & "','" & Replace(RS!ISS, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & Cartao & "')")
'                        End If
'
'                        DB.CommitTrans
'
'                    End If
'
'                RS.MoveNext
'            Loop
'            RS.Close
'    End If
'
'    '<<< RETIRADOS >>>
'    RS.Open "Select * From tblTitulo Where Data_Retirada BETWEEN'" & Format(Me.txtDataInicio, "mm/dd/yyyy") & "'AND'" & Format(Me.txtDataFim, "mm/dd/yyyy") & "' AND Tit_Particular='" & 1 & "' AND Anulado='" & 0 & "' AND Aguardando='" & 0 & "'", DB, adOpenDynamic
'    If RS.RecordCount = 0 Then
'        Me.cmdProcessando.Visible = False
'        Me.Refresh
'        RS.Close
'    Else
'            Do While Not RS.EOF
''                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Data_Retirada='" & Format(Me.txtData, "mm/dd/yyyy") & "' AND Baixa_Lote Is Null", DB, adOpenDynamic
'                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Data_Retirada BETWEEN'" & Format(Me.txtDataInicio, "mm/dd/yyyy") & "'AND'" & Format(Me.txtDataFim, "mm/dd/yyyy") & "'", DB, adOpenDynamic
'                If RSFinan.RecordCount = 0 Then
''                    Me.cmdProcessando.Visible = False
''                    Me.Refresh
''                    MsgBox "O Protocolo " & RS!Protocolo_Cartorio & " não foi lançado no Financeiro.", vbInformation
''                    RSFinan.Close
''                    RS.Close
''                    Exit Sub
'                End If
'
'                Do While Not RSFinan.EOF
'                    If IsNull(RSFinan!Estorno) Or RSFinan!Estorno = False Then
'                        DB.BeginTrans
'                        If RSFinan!Codigo = "1" Then
'                            TxBanco = 1
'                        Else
'                            TxBanco = 0
'                        End If
'
'                        If IsNull(RSFinan!FRJ) Then
'                            FRJ = 0
'                            FRC = 0
'                        Else
'                            FRJ = RSFinan!FRJ
'                            FRC = RSFinan!FRC
'                        End If
'
'                        If IsNull(RSFinan!Cartao) Or RSFinan!Cartao = False Then
'                            Cartao = 0
'                        Else
'                            Cartao = 1
'                        End If
'
'                        Valor_Distrib = RSFinan!Valor_Distrib - frjDistrib - frcDistrib
'                        V_Custas = RSFinan!Custas + frjDistrib + frcDistrib
'
'                        If RS!Tit_Particular = True Then
'                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Valor_Juros,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,Tipo_Baixa,TxBanco,ISS,FRJ,FRC,Cartao) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(V_Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & nNosso_Num & "','" & Replace(Portador, "'", " ") & "','" & Replace(RSFinan!Valor_Juros, ",", ".") & "','" & Replace(RSFinan!Valor_Mora, ",", ".") & "','" & Replace(RSFinan!V_Multa, ",", ".") & "','" & _
'                            Replace(RSFinan!Valor_CPMF, ",", ".") & "','" & RSFinan!Usuario & "','" & Replace(0, ",", ".") & "','" & RSFinan!Hora & "','" & "Retirados" & "','" & TxBanco & "','" & Replace(RSFinan!ISS, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & Cartao & "')")
'                        Else
'                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Valor_Juros,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,Tipo_Baixa,TxBanco,ISS,FRJ,FRC,Cartao) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(V_Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & nNosso_Num & "','" & Replace(Portador, "'", " ") & "','" & Replace(RSFinan!Valor_Juros, ",", ".") & "','" & Replace(RSFinan!Valor_Mora, ",", ".") & "','" & Replace(RSFinan!V_Multa, ",", ".") & "','" & _
'                            Replace(RSFinan!Valor_CPMF, ",", ".") & "','" & RSFinan!Usuario & "','" & Replace(0, ",", ".") & "','" & RSFinan!Hora & "','" & "Retirados" & "','" & TxBanco & "','" & Replace(RSFinan!ISS, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & Cartao & "')")
'                        End If
'                        DB.CommitTrans
'                        RSFinan.MoveNext
'                    Else
'                        RSFinan.MoveNext
'                    End If
'                Loop
'                RSFinan.Close
'                RS.MoveNext
'            Loop
'            RS.Close
'    End If
'
'
'    '<<< CUSTAS DISTRIBUIDOR >>>
''    SQL = "Select * From tblDepositoPrevio Where Data BETWEEN'" & Format(Me.txtDataInicio, "mm/dd/yyyy") & "'AND'" & Format(Me.txtDataFim, "mm/dd/yyyy") & "' AND Restituicao='" & 1 & "'"
''    RS.Open SQL, DB, adOpenDynamic
''    If RS.RecordCount = 0 Then
''        Me.cmdProcessando.Visible = False
''        Me.Refresh
''        RS.Close
''    Else
''        Do While Not RS.EOF
''            DB.BeginTrans
''            If RSFinan!Codigo = "1" Then
''                TxBanco = 1
''            Else
''                TxBanco = 0
''            End If
''
''            If IsNull(RSFinan!FRJ) Then
''                FRJ = 0
''                FRC = 0
''            Else
''                FRJ = RSFinan!FRJ
''                FRC = RSFinan!FRC
''            End If
''
''            If IsNull(RSFinan!Cartao) Or RSFinan!Cartao = False Then
''                Cartao = 0
''            Else
''                Cartao = 1
''            End If
''
''                DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Custas,Selo,Distrib,Devedor,Valor_Juros,Valor_Mora,V_Multa,CPMF,rDistrib,Tipo_Baixa,TxBanco,ISS,FRJ,FRC,Cartao) values ('" & RS!Protocolo & "','" & Replace(RS!Adiantamento, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & _
''                "','" & Replace(0, ",", ".") & "','" & "Custas Distribuidor" & "','" & TxBanco & "','" & Replace(RSFinan!ISS, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & Cartao & "')")
''            DB.CommitTrans
''            RS.MoveNext
''        Loop
''        RS.Close
''    End If
'
'    '<<< CUSTAS PROTESTO >>>
'    RS.Open "Select * From tblTitulo Where Data_Pagamento BETWEEN'" & Format(Me.txtDataInicio, "mm/dd/yyyy") & "'AND'" & Format(Me.txtDataFim, "mm/dd/yyyy") & "' AND CustasProtesto='" & 1 & "' AND Anulado='" & 0 & "' AND Aguardando='" & 0 & "' AND CancelaBanco='" & 0 & "'", DB, adOpenDynamic
'    If RS.RecordCount = 0 Then
'        Me.cmdProcessando.Visible = False
'        Me.Refresh
'        RS.Close
'    Else
'            Do While Not RS.EOF
'                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Data_Pagamento BETWEEN'" & Format(Me.txtDataInicio, "mm/dd/yyyy") & "'AND'" & Format(Me.txtDataFim, "mm/dd/yyyy") & "'AND Baixa_Lote Is Null", DB, adOpenDynamic
'                If RSFinan.RecordCount = 0 Then
''                    Me.cmdProcessando.Visible = False
''                    Me.Refresh
''                    MsgBox "O Protocolo " & RS!Protocolo_Cartorio & " não foi lançado no Financeiro.", vbInformation
''                    RSFinan.Close
''                    RS.Close
''                    Exit Sub
'                End If
'
'                Do While Not RSFinan.EOF
'                    If IsNull(RSFinan!Estorno) Or RSFinan!Estorno = False Then
'                        DB.BeginTrans
'                        If RSFinan!Codigo = "1" Then
'                            TxBanco = 1
'                        Else
'                            TxBanco = 0
'                        End If
'
'                        If IsNull(RSFinan!TaxaCartao) Then
'                            TaxaCartao = 0
'                        Else
'                            TaxaCartao = RSFinan!TaxaCartao
'                        End If
'
'                        If IsNull(RSFinan!FRJ) Then
'                            FRJ = 0
'                            FRC = 0
'                        Else
'                            FRJ = RSFinan!FRJ
'                            FRC = RSFinan!FRC
'                        End If
'
'                        If IsNull(RSFinan!Cartao) Or RSFinan!Cartao = False Then
'                            Cartao = 0
'                        Else
'                            Cartao = 1
'                        End If
'
'                        If IsNull(RSFinan!ISS) Then
'                            ISS = 0
'                        Else
'                            ISS = RSFinan!ISS
'                        End If
'
'                        Custas = RSFinan!Custas + frjDistrib + frcDistrib
'                        Valor_Distrib = RSFinan!Valor_Distrib - frjDistrib - frcDistrib
'
'                        DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Valor_Juros,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,Tipo_Baixa,TxBanco,ISS,FRJ,FRC,Cartao) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(Custas + TaxaCartao, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & nNosso_Num & "','" & Replace(Portador, "'", " ") & "','" & Replace(RSFinan!Valor_Juros, ",", ".") & "','" & Replace(RSFinan!Valor_Mora, ",", ".") & "','" & _
'                        Replace(RSFinan!V_Multa, ",", ".") & "','" & Replace(RSFinan!Valor_CPMF, ",", ".") & "','" & RSFinan!Usuario & "','" & Replace(0, ",", ".") & "','" & RSFinan!Hora & "','" & "Custas Protesto" & "','" & TxBanco & "','" & Replace(ISS, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & Cartao & "')")
'                        DB.CommitTrans
'                        RSFinan.MoveNext
'                    Else
'                        RSFinan.MoveNext
'                    End If
'                Loop
'                RSFinan.Close
'                RS.MoveNext
'            Loop
'            RS.Close
'    End If
'
'    '<<< DEPÓSITO PRÉVIO >>>
''    SQL = "Select * From tblDepositoPrevio Where Data BETWEEN'" & Format(Me.txtDataInicio, "mm/dd/yyyy") & "'AND'" & Format(Me.txtDataFim, "mm/dd/yyyy") & "' AND Restituicao='" & 0 & "'"
''    RS.Open SQL, DB, adOpenDynamic
''    If RS.RecordCount = 0 Then
''        Me.cmdProcessando.Visible = False
''        Me.Refresh
''        RS.Close
''    Else
''        Do While Not RS.EOF
''            DB.BeginTrans
'''            If RSFinan!Codigo = "1" Then
'''                TxBanco = 1
'''            Else
'''                TxBanco = 0
'''            End If
''                If IsNull(RSFinan!FRJ) Then
''                    FRJ = 0
''                    FRC = 0
''                Else
''                    FRJ = RSFinan!FRJ
''                    FRC = RSFinan!FRC
''                End If
''
''                If IsNull(RSFinan!Cartao) Or RSFinan!Cartao = False Then
''                    Cartao = 0
''                Else
''                    Cartao = 1
''                End If
''
''                DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Custas,Selo,Distrib,Devedor,Valor_Juros,Valor_Mora,V_Multa,CPMF,rDistrib,Tipo_Baixa,TxBanco,ISS,FRJ,FRC,Cartao) values ('" & RS!Protocolo & "','" & Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & _
''                "','" & Replace(RS!Adiantamento, ",", ".") & "','" & "Depósito Prévio" & "','" & TxBanco & "','" & Replace(RSFinan!ISS, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & Cartao & "')")
''            DB.CommitTrans
''            RS.MoveNext
''        Loop
''        RS.Close
''    End If
'
''<<< CERTIDÕES >>>
'    RSFinan.Open "SELECT DocCertidao, COALESCE(SUM(VlrCertidao), 0) AS VlrCertidao, COALESCE(SUM(Valor_Selo), 0) AS Valor_Selo, COALESCE(SUM(TaxaCartao), 0) AS TaxaCartao, COALESCE(SUM(tblFinanceiro.FRJ), 0) AS FRJ, COALESCE(SUM(tblFinanceiro.FRC), 0) AS FRC, MAX(CAST(COALESCE(Cartao, 0) AS INT)) AS Cartao, tblFinanceiro.Codigo, COALESCE(SUM(ISS), 0) AS ISS, CASE WHEN EXISTS (SELECT 1 FROM tblFinanceiro AS f2 WHERE f2.DocCertidao = tblFinanceiro.DocCertidao AND f2.TipoCertidao = 'POSITIVA' AND COALESCE(f2.VlrCertidao, 0) > 0) THEN 'POSITIVA' ELSE 'NEGATIVA' END AS TipoCertidao FROM tblFinanceiro INNER JOIN tblReqCertidao ON tblFinanceiro.DocCertidao = tblReqCertidao.Num_Doc WHERE Data_Certidao BETWEEN '" & Format(Me.txtDataInicio, "mm/dd/yyyy") & "' AND '" & Format(Me.txtDataFim, "mm/dd/yyyy") & "' AND VlrCertidao <> 0 AND tblReqCertidao.Pago = 1 GROUP BY DocCertidao, tblFinanceiro.Codigo", DB, adOpenDynamic
'    Do While Not RSFinan.EOF
'
'            If RSFinan!Codigo = "1" Then
'                TxBanco = 1
'            Else
'                TxBanco = Replace(TxBanco, ",", ".")
'            End If
'
'            DB.BeginTrans
'            DB.Execute ("Insert Into tblMovimentoCaixa (Custas,Selo,Distrib,Valor_Juros,Valor_Mora,V_Multa,CPMF,rDistrib,Tipo_Baixa,TxBanco,ISS,FRJ,FRC,Cartao) values ('" & Replace(RSFinan!VlrCertidao, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & "','" & _
'            Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & "','" & "Certidões" & "','" & TxBanco & "','" & Replace(RSFinan!ISS, ",", ".") & "','" & Replace(RSFinan!FRJ, ",", ".") & "','" & Replace(RSFinan!FRC, ",", ".") & "','" & RSFinan!Cartao & "')")
'            DB.CommitTrans
'            RSFinan.MoveNext
'    Loop
'    RSFinan.Close
'
''<<< IMPRESSÃO >>>
'            Me.cmdProcessando.Visible = False
'            Me.txtCarregando.Visible = True
'            Me.Refresh
'            PrintConnect
'            Rpt.Connect = Conexao
'            SQL = "Select * From tblMovimentoCaixa Order by Tipo_Baixa"
'            Screen.MousePointer = 11
'            Rpt.ReportFileName = App.Path & "\Fechamento.rpt"
'            Rpt.Destination = crptToWindow
'            Rpt.SQLQuery = SQL
'            Rpt.Formulas(0) = "Data='" & Me.txtDataInicio & " À " & Me.txtDataFim & "'"
'            Rpt.WindowState = crptMaximized
'            Rpt.CopiesToPrinter = 1
'            Me.txtCarregando.Visible = False
'            Rpt.Action = 1
'            Screen.MousePointer = 1
''<<< FIM IMPRESSÃO >>>
'            DB.Execute "Delete From tblMovimentoCaixa"
'Exit Sub
'Erro:
'    DB.RollbackTrans
'    MsgBox "Erro de Sistema. " & Err.Description, vbCritical
'    DB.Execute "Delete From tblMovimentoCaixa"
'    Me.cmdProcessando.Visible = False
'    Me.txtCarregando.Visible = False
'    DB.Execute "Delete From tblMovimentoCaixa"


End Sub


Private Sub cmdLimpaData_Click()
    Me.txtDataInicio = ""
    Me.txtDataFim = ""
    Me.cmdCancelados.Enabled = False
    Me.cmdCertidao.Enabled = False
    Me.cmdPagamentos.Enabled = False
    Me.cmdProtestados.Enabled = False
    Me.cmdRetirados.Enabled = False
    Me.cmdRetiradosPart.Enabled = False
    Me.cmdCustasProtesto.Enabled = False
    Me.cmdDepositoPrevio.Enabled = False
    Me.cmdCustas.Enabled = False
    Me.cmdFechamento.Enabled = False
    Me.cmdCDA.Enabled = False
    Me.cmdBoletos.Enabled = False

End Sub

Private Sub cmdPagamentos_Click()
On Error GoTo Erro
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim RSFinan As ADODB.Recordset
    Set RSFinan = New ADODB.Recordset
    Dim RSMovCaixa As ADODB.Recordset
    Set RSMovCaixa = New ADODB.Recordset
    Dim Aponta As ADODB.Recordset
    Set Aponta = New ADODB.Recordset
    Dim DataPag As Date
    
    Aponta.Open "Select * FROM tblApontamento", DB, adOpenDynamic
    DB.Execute "Delete From tblMovimentoCaixa"
    RSMovCaixa.Open "Select * From tblMovimentoCaixa WHERE TxBanco ='" & 0 & "'", DB, adOpenDynamic
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
    
    Me.cmdPagamentos.Enabled = False
    Me.cmdProcessando.Visible = True
    Me.Refresh
    RS.Open "Select * From tblTitulo Where Data_Pagamento Between '" & Format(Me.txtDataInicio, "mm/dd/yyyy") & "' AND '" & Format(Me.txtDataFim, "mm/dd/yyyy") & "' AND CustasProtesto Is Null AND Baixado='" & 1 & "' AND Anulado='" & 0 & "' AND CancelaBanco='" & 0 & "' AND Aguardando='" & 0 & "'", DB, adOpenDynamic
    If RS.RecordCount = 0 Then
        Me.cmdProcessando.Visible = False
        Me.Refresh
        MsgBox "Sem Dados para exibir!", vbInformation
        RS.Close
        Exit Sub
    Else
        'nResp = MsgBox("Confirma a impressão de " & RS.RecordCount & " Títulos Pagos?", vbQuestion + vbYesNo)
        'If nResp = vbYes Then
            Do While Not RS.EOF
'                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Data_Pagamento='" & Format(Me.txtData, "mm/dd/yyyy") & "'AND Baixa_Lote Is Null", DB, adOpenDynamic
                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Data_Pagamento Between '" & Format(Me.txtDataInicio, "mm/dd/yyyy") & "' AND '" & Format(Me.txtDataFim, "mm/dd/yyyy") & "'", DB, adOpenDynamic
                If RSFinan.RecordCount = 0 Then
'                    Me.cmdProcessando.Visible = False
'                    Me.Refresh
'                    MsgBox "O Protocolo " & RS!Protocolo_Cartorio & " não foi lançado no Financeiro.", vbInformation
'                    RSFinan.Close
'                    RS.Close
'                    Exit Sub
                End If
                                
                                
Calcula_Atos Aponta
                                
                Do While Not RSFinan.EOF
                    If IsNull(RSFinan!Estorno) Or RSFinan!Estorno = False Then
                        DB.BeginTrans
                        DataPag = "01/01/2020"
                        If IsNull(RSFinan!TED) Then
                            TED = 0
                        Else
                            TED = RSFinan!TED
                        End If
                        
                        If IsNull(RSFinan!TaxaCartao) Then
                            TxCartao = 0
                        Else
                            TxCartao = RSFinan!TaxaCartao
                        End If
                        
                        If IsNull(RSFinan!Cartao) Or RSFinan!Cartao = 0 Then
                            Cartao = 0
                        Else
                            Cartao = RSFinan!Cartao
                        End If
                        
                        If Format(RS!Data_Pagamento, "mm/dd/yyyy") >= DataPag Then
                            If IsNull(RSFinan!FRJ) Then
                                DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Valor_Juros,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,FRJ,FRC,TxBanco,TxCartao,ISS,Cartao) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(RSFinan!Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Replace(Portador, "'", " ") & "','" & Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & "','" & _
                                Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & "','" & RSFinan!Usuario & "','" & Replace(0, ",", ".") & "','" & RSFinan!Hora & "','" & Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(TED, ",", ".") & "','" & Replace(TxCartao, ",", ".") & "','" & Replace(RSFinan!ISS, ",", ".") & "','" & Cartao & "')")
                            Else
                                DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Valor_Juros,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,FRJ,FRC,TxBanco,TxCartao,ISS,Cartao) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(vDistrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Replace(Portador, "'", " ") & "','" & Replace(RSFinan!Valor_Juros, ",", ".") & "','" & Replace(RSFinan!Valor_Mora, ",", ".") & "','" & _
                                Replace(RSFinan!V_Multa, ",", ".") & "','" & Replace(RSFinan!Valor_CPMF, ",", ".") & "','" & RSFinan!Usuario & "','" & Replace(0, ",", ".") & "','" & RSFinan!Hora & "','" & Replace(RSFinan!FRJ - frjDistrib, ",", ".") & "','" & Replace(RSFinan!FRC - frcDistrib, ",", ".") & "','" & Replace(TED, ",", ".") & "','" & Replace(TxCartao, ",", ".") & "','" & Replace(RSFinan!ISS, ",", ".") & "','" & Cartao & "')")
                            End If
                        Else
                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Valor_Juros,Valor_Mora,V_Multa,CPMF,Usuario,rDistrib,Hora,FRJ,FRC,TxBanco,TxCartao,ISS,Cartao) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(RSFinan!Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Replace(Portador, "'", " ") & "','" & Replace(RSFinan!Valor_Juros, ",", ".") & "','" & Replace(RSFinan!Valor_Mora, ",", ".") & "','" & _
                            Replace(RSFinan!V_Multa, ",", ".") & "','" & Replace(RSFinan!Valor_CPMF, ",", ".") & "','" & RSFinan!Usuario & "','" & Replace(0, ",", ".") & "','" & RSFinan!Hora & "','" & Replace(0, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(TED, ",", ".") & "','" & Replace(TxCartao, ",", ".") & "','" & Replace(RSFinan!ISS, ",", ".") & "','" & Cartao & "')")
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
        'Else
        '    Me.cmdProcessando.Visible = False
        '    MsgBox "Impressão Cancelada!", vbInformation
        '    RS.Close
        '    DB.Execute "Delete From tblMovimentoCaixa"
        '    Exit Sub
        'End If

'<<< IMPRESSÃO >>>
            Me.cmdProcessando.Visible = False
            Me.txtCarregando.Visible = True
            Me.Refresh
            PrintConnect
            Rpt.Connect = Conexao
'            SQL = "Select * From tblMovimentoCaixa WHERE Cartao='" & 0 & "' Order by Protocolo"
            SQL = "Select * From tblMovimentoCaixa  Order by Protocolo"
            Screen.MousePointer = 11
            Rpt.ReportFileName = App.Path & "\Mov_Pagamento.rpt"
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
            Aponta.Close
    End If
Exit Sub
Erro:
    DB.RollbackTrans
    MsgBox "Erro de Sistema. " & Err.Description, vbCritical
    DB.Execute "Delete From tblMovimentoCaixa"
    Me.cmdProcessando.Visible = False
    Me.txtCarregando.Visible = False
    Aponta.Close
End Sub

Private Sub cmdProtestados_Click()
On Error GoTo Erro
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim RSFinan As ADODB.Recordset
    Set RSFinan = New ADODB.Recordset
    Dim RSMovCaixa As ADODB.Recordset
    Set RSMovCaixa = New ADODB.Recordset
    
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
    
    Me.cmdProtestados.Enabled = False
    Me.cmdProcessando.Visible = True
    Me.Refresh
    RS.Open "Select * From tblTitulo Where Data_Protestado='" & Format(Me.txtData, "mm/dd/yyyy") & "' AND Anulado='" & 0 & "' AND Aguardando='" & 0 & "'", DB, adOpenDynamic
    If RS.RecordCount = 0 Then
        Me.cmdProcessando.Visible = False
        Me.Refresh
        MsgBox "Sem Dados para exibir!", vbInformation
        RS.Close
        Exit Sub
    Else
        nResp = MsgBox("Confirma a impressão de " & RS.RecordCount & " Títulos Protestados?", vbQuestion + vbYesNo)
        If nResp = vbYes Then
            Do While Not RS.EOF
                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Data_Protesto='" & Format(Me.txtData, "mm/dd/yyyy") & "'", DB, adOpenDynamic
                If RSFinan.RecordCount = 0 Then
                    Me.cmdProcessando.Visible = False
                    Me.Refresh
                    MsgBox "O Protocolo " & RS!Protocolo_Cartorio & " não foi lançado no Financeiro.", vbInformation
                    RSFinan.Close
                    RS.Close
                    Exit Sub
                End If
                Do While Not RSFinan.EOF
                    If IsNull(RSFinan!Estorno) Or RSFinan!Estorno = False Then
                        DB.BeginTrans
                        'If Format(RS!Data_Apresenta, "mm/dd/yyyy") >= "12/09/2008" And Format(RS!Data_Apresenta, "mm/dd/yyyy") < "12/15/2008" Then
                        '    DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Livro,Folha) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "','" & RS!Livro & "','" & RS!Pagina & "')")
                        'Else
                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,Livro,Folha) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(RSFinan!Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "','" & RS!Livro & "','" & RS!Pagina & "')")
                        'End If
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
            Rpt.ReportFileName = App.Path & "\Mov_Protestado.rpt"
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
    End If
Exit Sub
Erro:
    DB.RollbackTrans
    MsgBox "Erro de Sistema. " & Err.Description, vbCritical
    DB.Execute "Delete From tblMovimentoCaixa"
    Me.cmdProcessando.Visible = False
    Me.txtCarregando.Visible = False
    

End Sub

Private Sub cmdRetirados_Click()
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
    
    Me.cmdRetirados.Enabled = False
    Me.cmdProcessando.Visible = True
    Me.Refresh
    RS.Open "Select * From tblTitulo Where Data_Retirada='" & Format(Me.txtData, "mm/dd/yyyy") & "' AND Tit_Particular='" & 0 & "' AND Anulado='" & 0 & "' AND Aguardando='" & 0 & "' AND CancelaBanco='" & 0 & "'", DB, adOpenDynamic
    If RS.RecordCount = 0 Then
        Me.cmdProcessando.Visible = False
        Me.Refresh
        MsgBox "Sem Dados para exibir!", vbInformation
        RS.Close
        Exit Sub
    Else
'        nResp = MsgBox("Confirma a impressão de " & RS.RecordCount & " Títulos Retirados?", vbQuestion + vbYesNo)
        nResp = vbYes
        If nResp = vbYes Then
            Do While Not RS.EOF
'                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Data_Retirada='" & Format(Me.txtData, "mm/dd/yyyy") & "'AND Baixa_Lote Is Null", DB, adOpenDynamic
                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Data_Retirada='" & Format(Me.txtData, "mm/dd/yyyy") & "'", DB, adOpenDynamic
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
                        'If Format(RS!Data_Apresenta, "mm/dd/yyyy") >= "12/09/2008" And Format(RS!Data_Apresenta, "mm/dd/yyyy") < "12/15/2008" Then
                        '    DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "')")
                        'Else
                        If Not IsNull(RSFinan!TaxaCartao) Then
                            vTaxa = RSFinan!TaxaCartao
                        Else
                            vTaxa = 0
                        End If
                            DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,ISS) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Custas + vTaxa, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(RSFinan!Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "','" & Replace(RSFinan!ISS, ",", ".") & "')")
                        'End If
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
            Rpt.ReportFileName = App.Path & "\Mov_Retirada.rpt"
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
    End If
Exit Sub
Erro:
    DB.RollbackTrans
    MsgBox "Erro de Sistema. " & Err.Description, vbCritical
    DB.Execute "Delete From tblMovimentoCaixa"
    Me.cmdProcessando.Visible = False
    Me.txtCarregando.Visible = False

End Sub



Private Sub cmdRetiradosPart_Click()
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
    
    Me.cmdRetiradosPart.Enabled = False
    Me.cmdProcessando.Visible = True
    Me.Refresh
    RS.Open "Select * From tblTitulo Where Data_Retirada='" & Format(Me.txtData, "mm/dd/yyyy") & "' AND Tit_Particular='" & 1 & "' AND Anulado='" & 0 & "' AND Aguardando='" & 0 & "' AND CancelaBanco='" & 0 & "'", DB, adOpenDynamic
    If RS.RecordCount = 0 Then
        Me.cmdProcessando.Visible = False
        Me.Refresh
        MsgBox "Sem Dados para exibir!", vbInformation
        RS.Close
        Exit Sub
    Else
'        nResp = MsgBox("Confirma a impressão de " & RS.RecordCount & " Títulos Retirados?", vbQuestion + vbYesNo)
        nResp = vbYes
        If nResp = vbYes Then
            Do While Not RS.EOF
                RSFinan.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Data_Retirada='" & Format(Me.txtData, "mm/dd/yyyy") & "'", DB, adOpenDynamic
                If RSFinan.RecordCount = 0 Then
                    Me.cmdProcessando.Visible = False
                    Me.Refresh
                    MsgBox "O Protocolo " & RS!Protocolo_Cartorio & " não foi lançado no Financeiro.", vbInformation
                    RSFinan.Close
                    RS.Close
                    Exit Sub
                End If
                Do While Not RSFinan.EOF
                    If IsNull(RSFinan!Estorno) Or RSFinan!Estorno = False Then
                        DB.BeginTrans
                        'If Format(RS!Data_Apresenta, "mm/dd/yyyy") >= "12/09/2008" And Format(RS!Data_Apresenta, "mm/dd/yyyy") < "12/15/2008" Then
                        '    DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Custas, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(0, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "')")
                        'Else
                        If Not IsNull(RSFinan!TaxaCartao) Then
                            vTaxa = RSFinan!TaxaCartao
                        Else
                            vTaxa = 0
                        End If
                        
                        DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,ISS) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RSFinan!Custas + vTaxa, ",", ".") & "','" & Replace(RSFinan!Valor_Selo, ",", ".") & "','" & Replace(RSFinan!Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "','" & Replace(RSFinan!ISS, ",", ".") & "')")
                        'End If
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
            Rpt.ReportFileName = App.Path & "\Mov_Retirada.rpt"
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
    End If
Exit Sub
Erro:
    DB.RollbackTrans
    MsgBox "Erro de Sistema. " & Err.Description, vbCritical
    DB.Execute "Delete From tblMovimentoCaixa"
    Me.cmdProcessando.Visible = False
    Me.txtCarregando.Visible = False

End Sub

Private Sub Form_Load()
    Me.Calendario.Value = Date
End Sub
Public Function Calcula_Atos(Aponta As ADODB.Recordset)
    eAponta = Aponta!Apontamento
    frcAponta = Format(eAponta * 0.025, "0.00")
    frjAponta = Format(eAponta * 0.15, "0.00")
    vAponta = eAponta + frcAponta + frjAponta
    
    
    eIntima = Aponta!Intimacao
    frcIntima = Format(eIntima * 0.025, "0.00")
    frjIntima = Format(eIntima * 0.15, "0.00")
    vIntima = eIntima + frcIntima + frjIntima
    
    eCanc = Aponta!CancAponta
    frcCanc = Format(eCanc * 0.025, "0.00")
    frjCanc = Format(eCanc * 0.15, "0.00")
    vCanc = eCanc + frcCanc + frjCanc
    
    eDistrib = Aponta!Distribuidor
    frcDistrib = Format(eDistrib * 0.025, "0.00")
    frjDistrib = Format(eDistrib * 0.15, "0.00")
    vDistrib = eDistrib + frcDistrib + frjDistrib
    
    eCpd = Aponta!CPD
    frcCpd = Format(eCpd * 0.025, "0.00")
    frjCpd = Format(eCpd * 0.15, "0.00")
    vCpd = eCpd + frcCpd + frjCpd
    
    eEdital = Aponta!V_Edital
    frcEdital = Format(eEdital * 0.025, "0.00")
    frjEdital = Format(eEdital * 0.15, "0.00")
    vEdital = eEdital + frcEdital + frjEdital
    
    eCtProt = Aponta!ContraProtesto
    frcCtProt = Format(eCtProt * 0.025, "0.00")
    frjCtProt = Format(eCtProt * 0.15, "0.00")
    vCtProt = eCtProt + frcCtProt + frjCtProt
    
End Function
