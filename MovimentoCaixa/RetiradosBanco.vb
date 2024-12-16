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


Erro:
    DB.RollbackTrans
    MsgBox "Erro de Sistema. " & Err.Description, vbCritical
    DB.Execute "Delete From tblMovimentoCaixa"
    Me.cmdProcessando.Visible = False
    Me.txtCarregando.Visible = False

End Sub



