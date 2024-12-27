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
'    RSFinan.Open "SELECT * FROM tblFinanceiro INNER JOIN tblReqCertidao ON tblFinanceiro.Codigo = tblReqCertidao.Codigo WHERE Data_Certidao Between '" & Format(Me.txtDataInicio, "yyyy/mm/dd") & "' AND '" & Format(Me.txtDataFim, "yyyy/mm/dd") & "' AND tblReqCertidao.Pago = 1" & "AND VlrCertidao != 0", DB, adOpenDynamic
    RSFinan.Open "SELECT * FROM tblFinanceiro WHERE (Data_Certidao BETWEEN '" & Format(Me.txtDataInicio, "yyyy/mm/dd") & "' AND '" & Format(Me.txtDataFim, "yyyy/mm/dd") & "' AND (Cenprot IS NULL OR Cenprot = 0)) OR (Rec_Data BETWEEN '" & Format(Me.txtDataInicio, "yyyy/mm/dd") & "' AND '" & Format(Me.txtDataFim, "yyyy/mm/dd") & "' AND Cenprot = 1)", DB, adOpenDynamic
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
                        DB.Execute "INSERT INTO tblMovimentoCaixa (Devedor, Num_Doc, Saldo, Selo, Usuario, Tipo_Certidao, Hora, TxBanco, FRJ, FRC, ISS, Cartao,Cenprot,Data_Certidao,Codigo) " & _
                        "VALUES ('" & Replace(RSFinan!NomeCertidao, "'", "''") & "','" & RSFinan!DocCertidao & "','" & Replace(vCertidao, ",", ".") & "','" & Replace(Valor_Selo, ",", ".") & "','" & RSFinan!Usuario & "','" & RSFinan!TipoCertidao & "','" & RSFinan!Hora & "','" & Replace(TxBanco, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & Replace(ISS, ",", ".") & "','" & Cartao & "','" & RSFinan!CENPROT & "','" & Format(RSFinan!Data_Certidao, "yyyy-mm-dd") & "','" & RSFinan!Codigo & "')"
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

