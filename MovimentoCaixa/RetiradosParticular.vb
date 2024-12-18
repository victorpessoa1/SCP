Private Sub cmdRetiradosPart_Click()
On Error GoTo Erro
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
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

'    RS.Open "SELECT t.*, f.* " & _
            "FROM tblFinanceiro AS f " & _
            "LEFT JOIN tblTitulo AS t ON t.Protocolo_Cartorio = f.Protocolo " & _
            "WHERE (t.Anulado = '0' AND t.CancelaBanco = '0'  AND t.Aguardando = '0' AND t.Tit_Particular = '1' " & _
            "AND t.Data_Retirada BETWEEN '" & Format(Me.txtDataInicio, "yyyy/mm/dd") & "' AND '" & Format(Me.txtDataFim, "yyyy/mm/dd") & "') " & _
            "ORDER BY f.Protocolo", DB, adOpenDynamic

RS.Open "SELECT t.*, f.* " & _
        "FROM tblFinanceiro AS f " & _
        "INNER JOIN tblTitulo AS t ON t.Protocolo_Cartorio = f.Protocolo " & _
        "WHERE (Anulado = '0' AND Aguardando = '0' AND Tit_Particular = '1' AND (Baixa_lote = '0' OR Baixa_lote IS NULL) AND " & _
        "t.Data_Retirada BETWEEN '" & Format(Me.txtDataInicio, "yyyy/mm/dd") & "' AND '" & Format(Me.txtDataFim, "yyyy/mm/dd") & "' AND " & _
        "f.Data_Retirada BETWEEN '" & Format(Me.txtDataInicio, "yyyy/mm/dd") & "' AND '" & Format(Me.txtDataFim, "yyyy/mm/dd") & "') " & _
        "OR (Anulado = '0' AND Aguardando = '0' AND Tit_Particular = '1' AND Baixa_lote = '1' AND f.Data_Retirada > 0 AND " & _
        "f.Rec_data BETWEEN '" & Format(Me.txtDataInicio, "yyyy/mm/dd") & "' AND '" & Format(Me.txtDataFim, "yyyy/mm/dd") & "') AND f.Estorno IS NULL " & _
        "ORDER BY f.Protocolo", DB, adOpenDynamic


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
                    If IsNull(RS!Estorno) Or RS!Estorno = False Then
                        DB.BeginTrans
                        If Not IsNull(RS!TaxaCartao) Then
                            vTaxa = RS!TaxaCartao
                        Else
                            vTaxa = 0
                        End If

                        DB.Execute ("Insert Into tblMovimentoCaixa (Protocolo,Num_Titulo,Especie_Tit,Vencimento,Data_Apresenta,Origem,Saldo,Custas,Selo,Distrib,Devedor,Sacador,Nosso_Num,Portador,ISS,Cartao) values ('" & RS!Protocolo_Cartorio & "','" & RS!Num_Titulo & "','" & nEspecie & "','" & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Mid(RS!Vencimento, 5, 4) & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "','" & RS!Origem & "','" & Replace(RS!Saldo, ",", ".") & "','" & Replace(RS!Custas + vTaxa, ",", ".") & "','" & Replace(RS!Valor_Selo, ",", ".") & "','" & Replace(RS!Valor_Distrib, ",", ".") & "','" & Replace(RS!Devedor, "'", " ") & "','" & Replace(RS!Sacador, "'", " ") & "','" & RS!Nosso_Num & "','" & Portador & "','" & Replace(RS!ISS, ",", ".") & "','" & RS!Cartao & "')")
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


