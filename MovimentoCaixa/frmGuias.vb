Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long
Private Const conSwNormal = 1

Private Declare Sub GenerateBMP Lib "quricol32.dll" Alias "GenerateBMPW" (ByVal FileName As Long, ByVal Text As Long, ByVal Margin As Long, ByVal Size As Long, ByVal Level As TErrorCorretion)
Private RSPrt As ADODB.Recordset
Private CrysApp As New CRAXDDRT.Application
Private CrysRep As New CRAXDDRT.Report

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
    
    Dim issAponta As Currency
    Dim issPago As Currency
    Dim issIntima As Currency
    Dim issCanc As Currency
    Dim issDistrib As Currency
    Dim issCpd As Currency
    Dim issProt As Currency
    Dim issRet As Currency
    Dim issCtProt As Currency
    Dim issEdital As Currency
    
    Dim bEmol As Currency
    Dim bFRJ As Currency
    Dim bFRC As Currency
    
    Dim atoPago As Integer
    Dim atoProt As Integer
    Dim atoCanc As Integer
    Dim n72Horas As String
    Dim nLivro As Integer
    Dim nPagina As Integer
    Dim nAvalista As Integer
    Dim nFaixa0 As Integer
    Dim nAtos As Integer
    Dim nAvalEdital As Integer
    Dim nAvalIntima As Integer
    Dim NumSelo1 As String
    Dim NumSelo2 As String
    Dim NumSelo3 As String
    Dim NumSelo4 As String
    Dim NumSelo5 As String
    Dim NumSelo6 As String
    Dim NumSelo7 As String
    Dim NumSelo8 As String
    Dim NumSelo9 As String
    Dim NumSelo10 As String
    Dim NumSelo11 As String
    Dim NumSelo12 As String
    Dim NumSelo13 As String
    Dim NumSelo14 As String
    Dim SerieGeral As String
    Dim SerieCertidao As String
    Dim Cod As String
    Dim Valor As Currency
    Dim Juros As Currency
    Dim Custas As Currency
    Dim Selo As Currency
    Dim CPMF As Currency
    Dim Distribuidor As Currency
    Dim Mora As Currency
    Dim Multa As Currency
    Dim Tipo_Baixa As String
    Dim Texto1 As String
    Dim Texto2 As String
    Dim nNome_Recibo As String
    Dim vCheque As Integer
    Dim SeloGeral As Integer
    Dim SeloCertidao As Integer
    Dim Cheque3 As Integer
    Dim ContraProtesto As Integer
    Dim Data_Entrada As Date
    Dim Custas0 As Integer
    Dim vTaxa As Currency
    Dim Anuencia As Boolean
    Dim Especie_Tit As String
    Dim nSeloGratuito As Integer
    Dim nBaixa As Integer
    Dim x As Integer
    Dim y As Integer
    Dim txtValorCustas As Currency
    Dim Total As Currency

Private Sub cboBaixa_Change()
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset

RS.Open "SELECT * FROM tblTitulo WHERE Protocolo_Cartorio ='" & Me.txtProtocolo & "'", DB, adOpenDynamic
    Tipo_Baixa = Me.cboBaixa
    Calcula_Baixa RS
End Sub

Private Sub cmdBaixar_Click()
On Error GoTo Erro

    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim RSBaixa As ADODB.Recordset
    Set RSBaixa = New ADODB.Recordset
    Dim RSAval As ADODB.Recordset
    Set RSAval = New ADODB.Recordset
    Dim RSBanco As ADODB.Recordset
    Set RSBanco = New ADODB.Recordset
    Dim RSAdianta As ADODB.Recordset
    Set RSAdianta = New ADODB.Recordset
    Dim nDataProt As Date
    Dim RSAponta As ADODB.Recordset
    Set RSAponta = New ADODB.Recordset
    Dim RSSeloDigital As ADODB.Recordset
    Set RSSeloDigital = New ADODB.Recordset
    Dim RSSeloGratuito As ADODB.Recordset
    Set RSSeloGratuito = New ADODB.Recordset
    Dim RSProtesto As ADODB.Recordset
    Set RSProtesto = New ADODB.Recordset
    Dim RSPostecipa As ADODB.Recordset
    Set RSPostecipa = New ADODB.Recordset
    Dim RSCod As ADODB.Recordset
    Set RSCod = New ADODB.Recordset
    Dim dOcorrencia As Date
    Dim RSImp As ADODB.Recordset
    Set RSImp = New ADODB.Recordset
    Dim Custas As Double
    Dim RSBanco1 As ADODB.Recordset
    Set RSBanco1 = New ADODB.Recordset
    Dim Portador As String
    
    FRJ = 0
    FRC = 0
    totFRJ = 0
    totFRC = 0
    Multa = 0
    CPMF = 0
    Juros = 0
    SeloGeral = 0
    SeloCertidao = 0
    Mora = 0
    Portador = ""
    CancelaBanco = 0
    Anuencia = False
    CodNota = 0
    N_Guia = ""
    
'<<< Verifica Pasta >>>
    If Len(Dir("\\" & Server & "\ProtestoSCP\Caixa\" & Replace(Date, "/", ""), vbDirectory) & "") > 0 Then
    Else
        MkDir "\\" & Server & "\ProtestoSCP\Caixa\" & Replace(Date, "/", "")
    End If
    
    RSBaixa.Open "SELECT * FROM tblGuias WHERE Marca='" & 1 & "'AND Baixado='" & 0 & "'", DB, adOpenDynamic
    If RSBaixa.RecordCount = 0 Then
        MsgBox "Sem títulos selecionados para LIQUIDAR.", vbInformation
        Exit Sub
    End If
    Tipo_Ocorrencia = RSBaixa!Tipo_Baixa
'<<< Abre Tabela de Apontamento >>>
    RSAponta.Open "Select * From tblApontamento", DB, adOpenDynamic
    
    Protocolo = ""
    
    If Tipo_Ocorrencia = "CANCELAMENTO" Then
    
    Do While Not RSBaixa.EOF
'<<< C A N C E L A M E N T O >>>

'<<< Busca o Protocolo na Tabela tblTitulo >>>
    RS.Open "Select * from tblTitulo where Protocolo_Cartorio=" & RSBaixa!Protocolo, DB, adOpenDynamic

    If RS.RecordCount = 0 Then
        MsgBox "O Título " & RSBaixa!Protocolo & " não pertence a este cartório.", vbInformation
        Exit Sub
    End If
    
        If RS!Tipo_Ocorrencia = "2" Then
        
            Valor = RS!Saldo
            Tipo_Baixa = "Cancelamento"
            nFaixa0 = 0
            CalculoFaixas RSAponta
            
'            If Valor <= RSAponta!Faixa0 Then ValorCustas = FormatCurrency(RSAponta!Cancelado0, 2): vPago = RSAponta!sPago0: atoCanc = 154: nFaixa0 = 0
'            If Valor <= RSAponta!Faixa1 And Valor > RSAponta!Faixa0 Then ValorCustas = FormatCurrency(RSAponta!Cancelado1, 2): vPago = RSAponta!sPago1: atoCanc = 155
'            If Valor <= RSAponta!Faixa2 And Valor > RSAponta!Faixa1 Then ValorCustas = FormatCurrency(RSAponta!Cancelado2, 2): vPago = RSAponta!sPago2: atoCanc = 156
'            If Valor <= RSAponta!Faixa3 And Valor > RSAponta!Faixa2 Then ValorCustas = FormatCurrency(RSAponta!Cancelado3, 2): vPago = RSAponta!sPago3: atoCanc = 157
'            If Valor <= RSAponta!Faixa4 And Valor > RSAponta!Faixa3 Then ValorCustas = FormatCurrency(RSAponta!Cancelado4, 2): vPago = RSAponta!sPago4: atoCanc = 158
'            If Valor <= RSAponta!Faixa5 And Valor > RSAponta!Faixa4 Then ValorCustas = FormatCurrency(RSAponta!Cancelado5, 2): vPago = RSAponta!sPago5: atoCanc = 159
'            If Valor <= RSAponta!Faixa6 And Valor > RSAponta!Faixa5 Then ValorCustas = FormatCurrency(RSAponta!Cancelado6, 2): vPago = RSAponta!sPago6: atoCanc = 160
'            If Valor > RSAponta!Faixa6 Then ValorCustas = FormatCurrency(RSAponta!Cancelado7, 2): vPago = RSAponta!sPago7: atoCanc = 161
                             
            ePago = vPago
            frcPago = Format(ePago * 0.025, "0.00")
            frjPago = Format(ePago * 0.15, "0.00")
            vPago = ePago + frcPago + frjPago
                             
            eRet = vRet
            frcRet = Format(eRet * 0.025, "0.00")
            frjRet = Format(eRet * 0.15, "0.00")
            vRet = eRet + frcRet + frjRet
            
            eAponta = RSAponta!Apontamento
            frcAponta = Format(eAponta * 0.025, "0.00")
            frjAponta = Format(eAponta * 0.15, "0.00")
            vAponta = eAponta + frcAponta + frjAponta
            
            
            eIntima = RSAponta!Intimacao
            frcIntima = Format(eIntima * 0.025, "0.00")
            frjIntima = Format(eIntima * 0.15, "0.00")
            vIntima = eIntima + frcIntima + frjIntima
            
            eCanc = RSAponta!CancAponta
            frcCanc = Format(eCanc * 0.025, "0.00")
            frjCanc = Format(eCanc * 0.15, "0.00")
            vCanc = eCanc + frcCanc + frjCanc
            
            eDistrib = RSAponta!Distribuidor
            frcDistrib = Format(eDistrib * 0.025, "0.00")
            frjDistrib = Format(eDistrib * 0.15, "0.00")
            vDistrib = eDistrib + frcDistrib + frjDistrib
            
            eCpd = RSAponta!CPD
            frcCpd = Format(eCpd * 0.025, "0.00")
            frjCpd = Format(eCpd * 0.15, "0.00")
            vCpd = eCpd + frcCpd + frjCpd
            
            eEdital = RSAponta!V_Edital
            frcEdital = Format(eEdital * 0.025, "0.00")
            frjEdital = Format(eEdital * 0.15, "0.00")
            vEdital = eEdital + frcEdital + frjEdital
            
            eCtProt = RSAponta!ContraProtesto
            frcCtProt = Format(eCtProt * 0.025, "0.00")
            frjCtProt = Format(eCtProt * 0.15, "0.00")
            vCtProt = eCtProt + frcCtProt + frjCtProt
            
            nTipoCancela = "A"
            Tipo_Baixa = "Cancelamento"
            Custas = ValorCustas - vCpd
            Selo = RSAponta!Selo * 2
'            Distribuidor = vDistrib
            Distribuidor = 0
            Total = Custas + Selo
            Juros = 0
            dOcorrencia = Date
            Adiantamento = FormatCurrency(0, 2)
                                    
            If RS!CodPortador > 0 Then
                RSBanco.Open "Select * From tblBanco Where idBanco='" & RS!CodPortador & "'", DB, adOpenDynamic
                Portador = RSBanco!Nome_Banco
                CodNota = RSBanco!CodNota
                Cobranca = RSBanco!Cobranca
                RSBanco.Close
            Else
                Portador = RS!Portador
            End If

'            If CodNota > 1 Then
'                RSSeloDigital.Open "Select * From tblSeloDigital Where Usado='" & 0 & "' AND Tipo='" & 102 & "' Order By IdSelo", DB, adOpenDynamic
'                Custas = 0
'                Total = 0
'                Selo = 0
'            Else
                RSSeloDigital.Open "Select * From tblSeloDigital Where Usado='" & 0 & "' AND Tipo='" & 101 & "' Order By IdSelo", DB, adOpenDynamic
'            End If
            
            If RSSeloDigital.RecordCount = 0 Then
                MsgBox "Sistema sem selos Digitais.", vbInformation, "Sem Selos"
                Screen.MousePointer = 1
                Exit Sub
            End If
            RSSeloDigital.MoveFirst
                    
            RSAval.Open "Select * From tblAvalista Where Protocolo_Cartorio='" & RSBaixa!Protocolo & "'", DB, adOpenDynamic
            nAvalista = RSAval.RecordCount
            
            Do While Not RSAval.EOF
                DB.Execute ("Update tblAvalista set Protestado='" & 0 & "' Where Protocolo_Cartorio='" & RSBaixa!Protocolo & "'")
                RSAval.MoveNext
            Loop
                    
                    
                    
    '<<< PREPARA A TABELA tblCaixa >>>
        totFRC = 0
        totFRJ = 0
    
        RSPostecipa.Open "Select * From tblSeloDigital Where Protocolo='" & RSBaixa!Protocolo & "' AND Tipo='" & 113 & "'", DB, adOpenDynamic
        If RSPostecipa.RecordCount > 0 Then
            yato = 2
            RSPostecipa.Close
        End If
            For xAto = 1 To 2
            '<<< Arredonda os centavos >>>
            Select Case xAto
                Case 1
                    Emolumento = vPago
                    Emol = ePago
                    CodAto = atoCanc
                    FRC = frcPago
                    FRJ = frjPago
                Case 2
                    Emolumento = vCanc
                    Emol = eCanc
                    CodAto = 893
                    FRC = frcCanc
                    FRJ = frjCanc
'                Case 3
'                    Emolumento = vDistrib
'                    Emol = eDistrib
'                    CodAto = 179
'                    FRC = frcDistrib
'                    FRJ = frjDistrib
            End Select
            If RSSeloDigital!Tipo = "102" Then
                FRC = 0
                FRJ = 0
                Emol = 0
            End If
            
            FRJ = Format(FRJ, "0.00")
            FRC = Format(FRC, "0.00")
            totFRJ = totFRJ + FRJ
            totFRC = totFRC + FRC
    
            Codigo = RSSeloDigital!Codigo
            Serie = RSSeloDigital!Serie
            Tipo = RSSeloDigital!Tipo
            CodSeguranca = RSSeloDigital!CodSeguranca
                
            Value = "https://apps.tjpa.jus.br/ValidaSeloDigital/Detalhes/Resultado?codigoSelo=" & Codigo & "&serie=" & Serie & "&codigoTipoSelo=" & Tipo & "&codigoSeguranca=" & CodSeguranca & ""
            Caminho = ("\\" & Server & "\ProtestoSCP\QRCimg\Cx") & Tipo_Baixa & RS!Protocolo_Cartorio & "(" & xAto & ")" & ".bmp"
            GenerateBMP StrPtr(Caminho), StrPtr(Value), 3, 5, QualityLow
            
            RSCod.Open "Select * From tblEspecieTitulos WHERE Especie='" & RS!Especie_Tit & "'", DB, adOpenDynamic
            NCod = RSCod!Codigo
            RSCod.Close
            
            DB.BeginTrans
            DB.Execute ("Insert Into tblCaixa (Protocolo_Cartorio,CodAto,QRCode,N_Selo,Serie,Cod_Seg,Emolumentos,FRJ,FRC,NCod,TipoSelo) values ('" & RS!Protocolo_Cartorio & "','" & CodAto & "','" & Caminho & "','" & Codigo & "','" & Serie & "','" & CodSeguranca & "','" & Replace(Emol, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & NCod & "','" & Tipo & "')")
            DB.Execute ("Update tblSeloDigital set Usado='" & 1 & "',Protocolo='" & RS!Protocolo_Cartorio & "',Data_Uso='" & Format(Date, "mm/dd/yyyy") & "',Ato='" & CodAto & "' Where IdSelo='" & RSSeloDigital!IdSelo & "'")
            DB.CommitTrans
            RSSeloDigital.MoveNext
            Next
    
                    If CDeb = 2 Then
                        Codigo = 1
                    Else
                        Codigo = 0
                    End If
    
                    DB.Execute ("Insert Into tblFinanceiro (CodPortador,Data_Ocorrencia,Protocolo,Custas,Tipo_Ocorrencia,Pagar,Valor_Tit,Valor_Juros,Valor_CPMF,Valor_Selo,Valor_Mora,V_Multa,Valor_Distrib,Data_Cancelado,Devedor,Usuario,Hora,CancelaBanco,QtdeSG,QtdeSC,Pagante,Especie_Tit,TaxaCartao,Codigo) values ('" & RS!CodPortador & "','" & Format(Date, "ddmmyyyy") & "','" & RS!Protocolo_Cartorio & "','" & Replace(Custas, ",", ".") & "','" & nTipoCancela & "','" & Replace(Total, ",", ".") & "','" & Replace(Valor, ",", ".") & "','" & Replace(Juros, ",", ".") & "','" & Replace(CPMF, ",", ".") & "','" & Replace(Selo, ",", ".") & "','" & Replace(Mora, ",", ".") & "','" & Replace(Multa, ",", ".") & "','" & 0 & "','" & Format(Date, "mm/dd/yyyy") & "','" & RS!Devedor & "','" & User & "','" & Time & "','" & CancelaBanco & "','" & SeloGeral & "','" & SeloCertidao & "','" & Portador & "','" & RS!Especie_Tit & "','" & Replace(vTaxa, ",", ".") & "','" & Codigo & "')")
'                    DB.Execute ("Update tblTitulo set Baixado='" & 1 & "',Tipo_Ocorrencia='" & nTipoCancela & "',Data_Ocorrencia='" & Format(Date, "ddmmyyyy") & "',Data_Cancelado='" & Format(Date, "mm/dd/yyyy") & "',CancelaBanco='" & CancelaBanco & "',Anuencia='" & Anuencia & "' Where Protocolo_Dist='" & RSBaixa!Protocolo & "'")
                    DB.Execute ("Update tblTitulo set Baixado='" & 1 & "',Tipo_Ocorrencia='" & nTipoCancela & "',Data_Ocorrencia='" & Format(Date, "ddmmyyyy") & "',Data_Cancelado='" & Format(Date, "mm/dd/yyyy") & "',CancelaBanco='" & CancelaBanco & "',Anuencia='" & Anuencia & "' Where Protocolo_Cartorio='" & RSBaixa!Protocolo & "'")
                    
                PrintConn
                
                Data_Ocorrencia = Format(Date, "ddmmyyyy")

                RSImp.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Tipo_Ocorrencia='" & "A" & "'AND Data_Ocorrencia='" & Data_Ocorrencia & "'AND Estorno Is Null", DB, adOpenDynamic
                If RSImp.RecordCount = 0 Then
                    MsgBox "Sem dados para impressão!", vbInformation
                    DB.RollbackTrans
                    Screen.MousePointer = 1
                    Exit Sub
                End If
                    
                Set RSPrt = New ADODB.Recordset
                If RSPrt.State = 1 Then RSPrt.Close
                RSPrt.Open "Select * From tblCaixa", Conn
                Set CrysRep = CrysApp.OpenReport(App.Path & "\Recibo_CaixaQRC.rpt")
'                FileLoca = ("\\" & Server & "\ProtestoSCP\Caixa\" & Replace(Date, "/", "") & "\Cancelamento" & RS!Protocolo_Cartorio & "_" & Replace(Time, ":", "") & ".pdf")
                FileLoca = ("\\" & Server & "\ProtestoSCP\Caixa\" & Replace(Date, "/", "") & "\" & UCase(Replace(Trim(RS!Devedor), "/", "")) & RS!Protocolo_Cartorio & "_CCL_" & Right(Replace(Time, ":", ""), 2) & ".pdf")
                If Not RSPrt.EOF Then
                    With CrysRep
                        Call .Database.Tables(1).SetDataSource(RSPrt)
                        If RS!CodPortador > 0 Then
                            RSBanco.Open "Select * From tblBanco Where idBanco='" & RS!CodPortador & "'", DB, adOpenDynamic
                            Portador = RSBanco!Nome_Banco
                            RSBanco.Close
                        Else
                            Portador = RS!Portador
                        End If
                        .DiscardSavedData
                        .EnableParameterPrompting = False
                        .ReadRecords
                        .ParameterFields(1).AddCurrentValue "Recibo de " & Tipo_Baixa & " " & FormatCurrency(RSImp!Pagar, 2)
                        .ParameterFields(2).AddCurrentValue "Protocolo: " & RSImp!Protocolo
                        .ParameterFields(3).AddCurrentValue "Recebemos de: " & UCase(Portador)
                        Texto = SeparaExtenso(Extenso(RSImp!Pagar), 40)
                        .ParameterFields(4).AddCurrentValue "A importância de: " & Texto1 & Texto2
                        .ParameterFields(5).AddCurrentValue "Referentes as custas de: " & Tipo_Baixa & ", do Apontamento e do Registro do Protesto"
                        .ParameterFields(6).AddCurrentValue "Do Titulo num. " & Trim(RS!Num_Titulo)
                        .ParameterFields(7).AddCurrentValue "Vencido em " & Format(RS!Vencimento, "00/00/0000")
                        .ParameterFields(8).AddCurrentValue "Sacador " & RS!Sacador
                        .ParameterFields(9).AddCurrentValue "Cedente " & RS!Cedente
                        .ParameterFields(10).AddCurrentValue "Apresentado por " & UCase(Portador)
                        .ParameterFields(11).AddCurrentValue "Contra :" & RS!Devedor & " Conforme discriminacao abaixo."
                        .ParameterFields(12).AddCurrentValue "Entrada " & RS!Data_Apresenta
                        .ParameterFields(13).AddCurrentValue "Nosso Numero " & RS!Nosso_Num
                        .ParameterFields(14).AddCurrentValue "Valor do Titulo: " & FormatCurrency(RSImp!Valor_Tit, 2)
                        If RSImp!Valor_Juros > 0 Then
                            .ParameterFields(15).AddCurrentValue "Juros: " & FormatCurrency(RSImp!Valor_Juros, 2)
                        End If
                                            
'                        nData = CDate("01/12/2019")
'                        If RS!Data_Apresenta < nData Then
                            .ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(totFRJ, 2)
'                        Else
'                            .ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(totFRJ - frjDistrib, 2)
'                            totFRJ = totFRJ - frjDistrib
'                        End If
'                        If RS!Data_Apresenta < nData Then
                            .ParameterFields(17).AddCurrentValue "Taxa FRC: " & FormatCurrency(totFRC, 2)
'                        Else
'                            .ParameterFields(17).AddCurrentValue "Taxa FRC: " & FormatCurrency(totFRC - frcDistrib, 2)
'                            totFRC = totFRC - frcDistrib
'                        End If
                        DB.BeginTrans
                        DB.Execute ("Update tblFinanceiro set FRJ='" & Replace(totFRJ, ",", ".") & "',FRC='" & Replace(totFRC, ",", ".") & "' Where idFinanceiro='" & RSImp!idFinanceiro & "'")
                        DB.CommitTrans
                        If RSImp!Valor_Selo > 0 Then
                            .ParameterFields(18).AddCurrentValue "Selo Judicial: " & FormatCurrency(RSImp!Valor_Selo, 2)
                        End If
'                        If RSImp!Valor_Distrib > 0 Then
'                            .ParameterFields(19).AddCurrentValue "Distribuidor: " & FormatCurrency(eDistrib, 2)
'                        End If
                        .ParameterFields(20).AddCurrentValue "Desp. Canc. Apont.: " & FormatCurrency(eCanc, 2)
                        'TT = FormatCurrency(RSImp!Custas - (RSAponta!CancAponta), 2)
                        .ParameterFields(21).AddCurrentValue "Desp. Cancelamento: " & FormatCurrency(ePago, 2)
                        If CDeb = 1 Then
                            .ParameterFields(22).AddCurrentValue "Desp. Cartão de Débito: " & FormatCurrency(RSImp!Pagar - RSImp!Custas - RSImp!Valor_Selo, 2)
                        End If
                        If CDeb = 2 Then
                            .ParameterFields(22).AddCurrentValue "Desp. Boleto Bancário: " & FormatCurrency(RSImp!TaxaCartao, 2)
                        End If
                        .ParameterFields(23).AddCurrentValue ""
                        'If Me.chkEditado = True Then
                            .ParameterFields(24).AddCurrentValue ""
                        '    TT = FormatCurrency(RSImp!Custas - (RSAponta!Apontamento + RSAponta!Intimacao + RSAponta!CancAponta + RSAponta!CPD + RSAponta!V_Edital), 2)
                        'Else
                        '    TT = FormatCurrency(RSImp!Custas - (RSAponta!CancAponta), 2)
                        'End If
                        .ParameterFields(25).AddCurrentValue ""
                        .ParameterFields(26).AddCurrentValue "TOTAL PAGO       :" & FormatCurrency(RSImp!Pagar, 2)
                        .ParameterFields(27).AddCurrentValue "" & "Belém, " & Day(CDate(Format(Data_Ocorrencia, "00/00/0000"))) & " de " & MonthName(Month(CDate(Format(Data_Ocorrencia, "00/00/0000")))) & " de " & Year(CDate(Format(Data_Ocorrencia, "00/00/0000")))
                        .ParameterFields(28).AddCurrentValue "Protocolo Dist.: " & RS!Protocolo_Dist
    '                    .ParameterFields(28).AddCurrentValue "Autenticacao: " & RSImp!Protocolo * (Day(Date) & Month(Date) & Year(Date))
                        '.ParameterFields(29).AddCurrentValue "" & NumSelo1 & "-" & NumSelo2 & "-" & NumSelo3 & "-" & NumSelo4 & "-" & NumSelo5 & "-" & NumSelo6 & "-" & NumSelo7
                    End With
                    Set CRExportOptions = CrysRep.ExportOptions
                    CRExportOptions.FormatType = crEFTPortableDocFormat
                    CRExportOptions.DestinationType = crEDTDiskFile
                    CRExportOptions.DiskFileName = FileLoca
                    CrysRep.DisplayProgressDialog = False
                    CrysRep.Export False
                    Set CRExportOptions = Nothing
                End If
                
                nResp = MsgBox("Imprimir o Recibo", vbQuestion + vbYesNo)
                If nResp = vbYes Then
                    CrysRep.PrintOut False, 1
                End If
                Screen.MousePointer = 1
                'frmPrtCaixa.CRViewer1.ReportSource = CrysRep
                'frmPrtCaixa.CRViewer1.ViewReport
                'frmPrtCaixa.Show 1
                If Conn.State = 1 Then Conn.Close
                CaixaXML
'                EnvioXML

            '<<< Apaga QRCode >>>
                RSPrt.Open "Select * From tblCaixa", DB, adOpenDynamic
                    Do While Not RSPrt.EOF
                        Kill RSPrt!QRCode
                        RSPrt.MoveNext
                    Loop
                RSPrt.Close

                DB.Execute "Delete From tblCaixa"

'<<< F I M   C A N C E L A M E N T O >>>

'<<< C U S T A S  D E  P R O T E S T O >>>
    If CodNota > 1 Then
    Else
    Valor = RS!Saldo
    CalculoFaixaProtesto RSAponta
'        If Valor <= RSAponta!Faixa0 Then ValorCustas = FormatCurrency(RSAponta!Pago720, 2): vPago = RSAponta!sPago0: atoPago = 171: vProt = RSAponta!sProt0: atoCanc = 154: atoProt = 144: nFaixa0 = 0
'        If Valor <= RSAponta!Faixa1 And Valor > RSAponta!Faixa0 Then ValorCustas = FormatCurrency(RSAponta!Pago721, 2): vPago = RSAponta!sPago1: atoPago = 172: vProt = RSAponta!sProt1: atoCanc = 155: atoProt = 145
'        If Valor <= RSAponta!Faixa2 And Valor > RSAponta!Faixa1 Then ValorCustas = FormatCurrency(RSAponta!Pago722, 2): vPago = RSAponta!sPago2: atoPago = 173: vProt = RSAponta!sProt2: atoCanc = 156: atoProt = 146
'        If Valor <= RSAponta!Faixa3 And Valor > RSAponta!Faixa2 Then ValorCustas = FormatCurrency(RSAponta!Pago723, 2): vPago = RSAponta!sPago3: atoPago = 174: vProt = RSAponta!sProt3: atoCanc = 157: atoProt = 147
'        If Valor <= RSAponta!Faixa4 And Valor > RSAponta!Faixa3 Then ValorCustas = FormatCurrency(RSAponta!Pago724, 2): vPago = RSAponta!sPago4: atoPago = 175: vProt = RSAponta!sProt4: atoCanc = 158: atoProt = 148
'        If Valor <= RSAponta!Faixa5 And Valor > RSAponta!Faixa4 Then ValorCustas = FormatCurrency(RSAponta!Pago725, 2): vPago = RSAponta!sPago5: atoPago = 176: vProt = RSAponta!sProt5: atoCanc = 159: atoProt = 149
'        If Valor <= RSAponta!Faixa6 And Valor > RSAponta!Faixa5 Then ValorCustas = FormatCurrency(RSAponta!Pago726, 2): vPago = RSAponta!sPago6: atoPago = 177: vProt = RSAponta!sProt6: atoCanc = 160: atoProt = 150
'        If Valor > RSAponta!Faixa6 Then ValorCustas = FormatCurrency(RSAponta!Pago727, 2): vPago = RSAponta!sPago7: atoPago = 178: vProt = RSAponta!sProt7: atoCanc = 161: atoProt = 151

        ePago = vPago
        frcPago = Format(ePago * 0.025, "0.00")
        frjPago = Format(ePago * 0.15, "0.00")
        vPago = ePago + frcPago + frjPago
                
        eProt = vProt
        frcProt = Format(eProt * 0.025, "0.00")
        frjProt = Format(eProt * 0.15, "0.00")
        vProt = eProt + frcProt + frjProt
        
        If RS!Editado = True Then
            nAtos = 5
            Custas = FormatCurrency(vProt + vAponta + vCpd + vIntima + vEdital, 2)
        Else
            nAtos = 4
            Custas = FormatCurrency(vProt + vAponta + vCpd + vIntima, 2)
        End If
        
        If RS!ContraProtesto = True And RS!Editado = True Then
            nAtos = 6
            ContraProtesto = 1
            Custas = FormatCurrency(vProt + vAponta + vCpd + vIntima + vEdital + vCtProt, 2)
        End If
        
        If RS!ContraProtesto = True And RS!Editado = False Then
            nAtos = 5
            ContraProtesto = 1
            Custas = FormatCurrency(vProt + vAponta + vCpd + vIntima + vCtProt, 2)
        End If
    
    'Verifica se tem Avalista
    RSAval.Close
    RSAval.Open "Select * From tblAvalista Where Protocolo_Cartorio='" & RS!Protocolo_Cartorio & "'", DB, adOpenDynamic
    
    Do While Not RSAval.EOF
        If RSAval!Editado = True Then
            nAvalEdital = nAvalEdital + 1
            Custas = Custas + vEdital
            nAtos = nAtos + 1
        End If
        If RSAval!Intimado = True Then
            nAvalIntima = nAvalIntima + 1
            Custas = Custas + vIntima
            nAtos = nAtos + 1
        End If
        nAvalista = nAvalista + 1
        RSAval.MoveNext
    Loop
    RSAval.Close
    'FIM Verifica se tem Avalista
    
    Selo = (nAtos + 1) * RSAponta!Selo
    Distribuidor = vDistrib
    Total = Custas + Selo + Distribuidor
    Tipo_Baixa = "Protesto"
    
    RSAdianta.Open "Select * From tblAdiantamento Where Protocolo_Cartorio='" & RS!Protocolo_Cartorio & "'", DB, adOpenDynamic
    If IsNull(RS!Adiantamento) Then
        Adiantamento = 0
    Else
        Adiantamento = RS!Adiantamento
    End If
    If RSAdianta.RecordCount = 0 Then
        'Adiantamento = FormatCurrency(RS!Adiantamento, 2)
        DB.Execute ("Insert Into tblAdiantamento (Protocolo_Cartorio,Devedor,Valor,Saldo,Sacador,Cedente,Adiantamento,Baixa,DataPagamento,Data_Entrada) values ('" & RS!Protocolo_Cartorio & "','" & RS!Devedor & "','" & Replace(RS!Valor, ",", ".") & "','" & Replace(RS!Saldo, ",", ".") & "','" & RS!Sacador & "','" & RS!Cedente & "','" & Replace(Adiantamento, ",", ".") & "','" & 1 & "','" & Format(Date, "mm/dd/yyyy") & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "')")
    Else
        'Adiantamento = FormatCurrency(RSAdianta!Adiantamento, 2)
        DB.Execute ("Update tblAdiantamento set Baixa='" & 1 & "',DataPagamento='" & Format(Date, "mm/dd/yyyy") & "' Where Protocolo_Cartorio='" & RS!Protocolo_Cartorio & "'")
    End If
    
    If RS!Fora_Area = True Then
        Custas = Custas
    Else
        Custas = Custas - Adiantamento
    End If
                    
    Codigo = 0
    
'    If CodNota > 0 Then
'        DB.Execute ("Update tblTitulo set CustasProtesto='" & 1 & "' Where Protocolo_Cartorio='" & RS!Protocolo_Cartorio & "'")
''        DB.Execute ("Update tblTitulo set CustasProtesto='" & 1 & "' Where Protocolo_Dist='" & RS!Protocolo_Cartorio & "'")
'        DB.Execute ("Insert Into tblFinanceiro (CodPortador,Data_Ocorrencia,Protocolo,Custas,Tipo_Ocorrencia,Pagar,Valor_Tit,Valor_Juros,Valor_CPMF,Valor_Selo,Valor_Mora,V_Multa,Valor_Distrib,Devedor,Usuario,Hora,QtdeSG,QtdeSC,Pgto_CH,Pagante,Especie_Tit,TaxaCartao,Codigo) values ('" & RS!CodPortador & "','" & Format(Date, "ddmmyyyy") & "','" & RS!Protocolo_Cartorio & "','" & Replace(Custas, ",", ".") & "','" & 1 & "','" & Replace(Total, ",", ".") & "','" & Replace(Valor, ",", ".") & "','" & Replace(Juros, ",", ".") & "','" & Replace(CPMF, ",", ".") & "','" & Replace(Selo, ",", ".") & "','" & Replace(Mora, ",", ".") & "','" & Replace(Multa, ",", ".") & "','" & Replace(Distribuidor, ",", ".") & "','" & RS!Devedor & "','" & User & "','" & Time & "','" & SeloGeral & "','" & SeloCertidao & "','" & vCheque & "','" & Portador & "','" & RS!Especie_Tit & "','" & Replace(vTaxa, ",", ".") & "','" & Codigo & "')")
'    Else
        DB.Execute ("Update tblTitulo set Data_Pagamento='" & Format(Date, "mm/dd/yyyy") & "',CustasProtesto='" & 1 & "' Where Protocolo_Cartorio='" & RS!Protocolo_Cartorio & "'")
'        DB.Execute ("Update tblTitulo set Data_Pagamento='" & Format(Date, "mm/dd/yyyy") & "',CustasProtesto='" & 1 & "' Where Protocolo_Dist='" & RS!Protocolo_Cartorio & "'")
        DB.Execute ("Insert Into tblFinanceiro (CodPortador,Data_Ocorrencia,Protocolo,Custas,Tipo_Ocorrencia,Pagar,Valor_Tit,Valor_Juros,Valor_CPMF,Valor_Selo,Valor_Mora,V_Multa,Valor_Distrib,Data_Pagamento,Devedor,Usuario,Hora,QtdeSG,QtdeSC,Pgto_CH,Pagante,Especie_Tit,TaxaCartao,Codigo) values ('" & RS!CodPortador & "','" & Format(Date, "ddmmyyyy") & "','" & RS!Protocolo_Cartorio & "','" & Replace(Custas, ",", ".") & "','" & 1 & "','" & Replace(Total, ",", ".") & "','" & Replace(Valor, ",", ".") & "','" & Replace(Juros, ",", ".") & "','" & Replace(CPMF, ",", ".") & "','" & Replace(Selo, ",", ".") & "','" & Replace(Mora, ",", ".") & "','" & Replace(Multa, ",", ".") & "','" & Replace(Distribuidor, ",", ".") & "','" & Format(Date, "mm/dd/yyyy") & "','" & RS!Devedor & "','" & User & "','" & Time & "','" & SeloGeral & "','" & SeloCertidao & "','" & 0 & "','" & Portador & "','" & RS!Especie_Tit & "','" & Replace(0, ",", ".") & "','" & Codigo & "')")
'    End If

'<<< PREPARA A TABELA tblCaixa >>>
'    If Not IsNull(RS!CodPortador) Then
'        If RS!CodPortador = 0 Then
'            Portador = RS!Portador
'        Else
'            RSBanco.Open "Select * from tblBanco Where idBanco = '" & RS!CodPortador & "'", DB, adOpenDynamic
'            Portador = RSBanco!Nome_Banco
'            RSBanco.Close
'        End If
'    End If
        
    yato = 1
    If RS!ContraProtesto = True Then
        yato = 2
        ContraProtesto = 1
    End If
    
    For xAto = 1 To yato
    '<<< Arredonda os centavos >>>
            
    Select Case xAto
        Case 1
            Emolumento = vCpd
            Emol = eCpd
            CodAto = 967
            FRC = frcCpd
            FRJ = frjCpd
        Case 2
            Emolumento = vCtProt
            Emol = eCtProt
            CodAto = 966
            FRC = frcCtProt
            FRJ = frjCtProt
    End Select
    
    FRJ = Format(FRJ, "0.00")
    FRC = Format(FRC, "0.00")
    totFRJ = totFRJ + FRJ
    totFRC = totFRC + FRC

    Codigo = RSSeloDigital!Codigo
    Serie = RSSeloDigital!Serie
    Tipo = RSSeloDigital!Tipo
    CodSeguranca = RSSeloDigital!CodSeguranca
        
    Value = "https://apps.tjpa.jus.br/ValidaSeloDigital/Detalhes/Resultado?codigoSelo=" & Codigo & "&serie=" & Serie & "&codigoTipoSelo=" & Tipo & "&codigoSeguranca=" & CodSeguranca & ""
    Caminho = ("\\" & Server & "\ProtestoSCP\QRCimg\Cx") & Tipo_Baixa & RS!Protocolo_Cartorio & "(" & xAto & ")" & ".bmp"
    GenerateBMP StrPtr(Caminho), StrPtr(Value), 3, 5, QualityLow
    
    RSCod.Open "Select * From tblEspecieTitulos WHERE Especie='" & RS!Especie_Tit & "'", DB, adOpenDynamic
    NCod = RSCod!Codigo
    RSCod.Close
    
    DB.BeginTrans
    DB.Execute ("Insert Into tblCaixa (Protocolo_Cartorio,CodAto,QRCode,N_Selo,Serie,Cod_Seg,Emolumentos,FRJ,FRC,NCod,TipoSelo) values ('" & RS!Protocolo_Cartorio & "','" & CodAto & "','" & Caminho & "','" & Codigo & "','" & Serie & "','" & CodSeguranca & "','" & Replace(Emol, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & NCod & "','" & Tipo & "')")
    DB.Execute ("Update tblSeloDigital set Usado='" & 1 & "',Protocolo='" & RS!Protocolo_Cartorio & "',Data_Uso='" & Format(Date, "mm/dd/yyyy") & "',Ato='" & CodAto & "' Where IdSelo='" & RSSeloDigital!IdSelo & "'")
    DB.CommitTrans
    RSSeloDigital.MoveNext
    Next
        
'<<< ROTINA DE VERIFICAÇÃO DOS SELOS POSTECIPADOS >>>
    If RSPostecipa.State = 1 Then RSPostecipa.Close
    RSPostecipa.Open "Select * From tblSeloDigital Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Tipo='" & 113 & "' AND Postecipado='" & 0 & "'", DB, adOpenDynamic
    If RSPostecipa.RecordCount > 0 Then
    
        yato = RSPostecipa.RecordCount
    
        For xAto = 1 To yato
        '<<< Arredonda os centavos >>>
        Salva = 0
        Codigo = RSPostecipa!Codigo
        Serie = RSPostecipa!Serie
        Tipo = RSPostecipa!Tipo
        CodSeguranca = RSPostecipa!CodSeguranca
        'Protesto
        If RSPostecipa!Ato >= 144 And RSPostecipa!Ato <= 151 Or RSPostecipa!Ato >= 827 And RSPostecipa!Ato <= 892 Then
            Emolumento = vProt
            Emol = eProt
            CodAto = RSPostecipa!Ato
            FRC = frcProt
            FRJ = frjProt
            Codigo = RSPostecipa!Codigo
            Serie = RSPostecipa!Serie
            Tipo = RSPostecipa!Tipo
            CodSeguranca = RSPostecipa!CodSeguranca
            Salva = 1
        End If
        'Apontamento
        If RSPostecipa!Ato = 152 Or RSPostecipa!Ato = 756 Then
            Emolumento = vAponta
            Emol = eAponta
            CodAto = RSPostecipa!Ato
            FRC = frcAponta
            FRJ = frjAponta
            Codigo = RSPostecipa!Codigo
            Serie = RSPostecipa!Serie
            Tipo = RSPostecipa!Tipo
            CodSeguranca = RSPostecipa!CodSeguranca
            Salva = 1
        End If
        'Distribuidor
        If RSPostecipa!Ato = 179 Or RSPostecipa!Ato = 755 Then
            Emolumento = vDistrib
            Emol = eDistrib
            CodAto = RSPostecipa!Ato
            FRC = frcDistrib
            FRJ = frjDistrib
            Codigo = RSPostecipa!Codigo
            Serie = RSPostecipa!Serie
            Tipo = RSPostecipa!Tipo
            CodSeguranca = RSPostecipa!CodSeguranca
            Salva = 1
        End If
        'Intimação
        If RSPostecipa!Ato = 162 Or RSPostecipa!Ato = 757 Then
            Emolumento = vIntima
            Emol = eIntima
            CodAto = RSPostecipa!Ato
            FRC = frcIntima
            FRJ = frjIntima
            Codigo = RSPostecipa!Codigo
            Serie = RSPostecipa!Serie
            Tipo = RSPostecipa!Tipo
            CodSeguranca = RSPostecipa!CodSeguranca
            Salva = 1
        End If
        'Edital
        If RSPostecipa!Ato = 164 Or RSPostecipa!Ato = 760 Then
            Emolumento = vEdital
            Emol = eEdital
            CodAto = RSPostecipa!Ato
            FRC = frcEdital
            FRJ = frjEdital
            Codigo = RSPostecipa!Codigo
            Serie = RSPostecipa!Serie
            Tipo = RSPostecipa!Tipo
            CodSeguranca = RSPostecipa!CodSeguranca
            Salva = 1
        End If
        
        FRJ = Format(FRJ, "0.00")
        FRC = Format(FRC, "0.00")
        totFRJ = totFRJ + FRJ
        totFRC = totFRC + FRC


        Value = "https://apps.tjpa.jus.br/ValidaSeloDigital/Detalhes/Resultado?codigoSelo=" & Codigo & "&serie=" & Serie & "&codigoTipoSelo=" & Tipo & "&codigoSeguranca=" & CodSeguranca & ""
        Caminho = ("\\" & Server & "\ProtestoSCP\QRCimg\Cx") & Tipo_Baixa & RS!Protocolo_Cartorio & "(" & xAto + 2 & ")" & ".bmp"
        GenerateBMP StrPtr(Caminho), StrPtr(Value), 3, 5, QualityLow
        
        RSCod.Open "Select * From tblEspecieTitulos WHERE Especie='" & RS!Especie_Tit & "'", DB, adOpenDynamic
        NCod = RSCod!Codigo
        RSCod.Close
        
        If Salva = 1 Then
            DB.BeginTrans
            DB.Execute ("Insert Into tblCaixa (Protocolo_Cartorio,CodAto,QRCode,N_Selo,Serie,Cod_Seg,Emolumentos,FRJ,FRC,NCod,TipoSelo) values ('" & RS!Protocolo_Cartorio & "','" & CodAto & "','" & Caminho & "','" & Codigo & "','" & Serie & "','" & CodSeguranca & "','" & Replace(Emol, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & NCod & "','" & Tipo & "')")
            DB.Execute ("Update tblSeloDigital set Postecipado='" & 1 & "',Data_Uso='" & Format(Date, "mm/dd/yyyy") & "' Where IdSelo='" & RSPostecipa!IdSelo & "'")
            DB.CommitTrans
            RSPostecipa.MoveNext
        End If
        Next

    End If
'<<< FIM PREPARA A TABELA tblCaixa >>>
        
     PrintConn
     If RSImp.State = 1 Then RSImp.Close
     RSImp.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Tipo_Ocorrencia='" & 1 & "'AND Data_Ocorrencia='" & Data_Ocorrencia & "'AND Estorno Is Null", DB, adOpenDynamic
     If RSImp.RecordCount = 0 Then
         MsgBox "Sem dados para impressão!", vbInformation
         DB.RollbackTrans
         Screen.MousePointer = 1
         RSImp.Close
         Exit Sub
     End If
                
     Set RSPrt = New ADODB.Recordset
     If RSPrt.State = 1 Then RSPrt.Close
     RSProtesto.Open "Select * From tblProtestoCopia Where Protocolo_Cartorio='" & RS!Protocolo_Cartorio & "'", DB, adOpenDynamic
     RSPrt.Open "Select * From tblCaixa", Conn
     Set CrysRep = CrysApp.OpenReport(App.Path & "\Recibo_CaixaQRC.rpt")
'     FileLoca = ("\\" & Server & "\ProtestoSCP\Caixa\" & Replace(Date, "/", "") & "\CustaProtesto" & RS!Protocolo_Cartorio & "_" & Replace(Time, ":", "") & ".pdf")
     FileLoca = ("\\" & Server & "\ProtestoSCP\Caixa\" & Replace(Date, "/", "") & "\" & UCase(Trim(Replace(RS!Devedor, "/", ""))) & RS!Protocolo_Cartorio & "_CPT_" & Right(Replace(Time, ":", ""), 2) & ".pdf")
     If Not RSPrt.EOF Then
         With CrysRep
             Call .Database.Tables(1).SetDataSource(RSPrt)
             .DiscardSavedData
             .EnableParameterPrompting = False
             .ReadRecords
             If RSAdianta.RecordCount = 0 Then
                 vTotal = FormatCurrency(RSImp!Pagar, 2) - FormatCurrency(Adiantamento, 2)
             Else
                 vTotal = FormatCurrency(RSImp!Pagar, 2) - FormatCurrency(RSAdianta!Adiantamento, 2)
             End If
             .ParameterFields(1).AddCurrentValue "Recibo de " & Tipo_Baixa & " " & FormatCurrency(RSImp!Pagar, 2)
             .ParameterFields(2).AddCurrentValue "Protocolo: " & RSImp!Protocolo
             .ParameterFields(3).AddCurrentValue "Recebemos de: " & UCase(Portador)
             Texto = SeparaExtenso(Extenso(RSImp!Pagar), 40)
             .ParameterFields(4).AddCurrentValue "A importância de: " & Texto1 & Texto2
             .ParameterFields(5).AddCurrentValue "Referentes as custas de Certidão de Protesto "
             .ParameterFields(6).AddCurrentValue "Do Titulo num. " & RS!Num_Titulo
             .ParameterFields(7).AddCurrentValue "Vencido em " & Format(RS!Vencimento, "00/00/0000")
             .ParameterFields(8).AddCurrentValue "Sacador " & RS!Sacador
             .ParameterFields(9).AddCurrentValue "Cedente " & RS!Cedente
             If RS!CodPortador > 0 Then
                 RSBanco.Open "Select * From tblBanco Where idBanco='" & RS!CodPortador & "'", DB, adOpenDynamic
                 .ParameterFields(10).AddCurrentValue "Apresentado por " & RSBanco!Nome_Banco
                 RSBanco.Close
             Else
                 .ParameterFields(10).AddCurrentValue "Apresentado por " & RS!Portador
             End If
             .ParameterFields(11).AddCurrentValue "Contra " & RS!Devedor & " Conforme discriminação abaixo."
             .ParameterFields(12).AddCurrentValue "Entrada " & RS!Data_Apresenta
             .ParameterFields(13).AddCurrentValue "Nosso Numero " & RS!Nosso_Num
             .ParameterFields(14).AddCurrentValue "Valor do Titulo   :" & FormatCurrency(RSImp!Valor_Tit, 2)
             If RSImp!Valor_Juros > 0 Then
                 .ParameterFields(15).AddCurrentValue "Juros             :" & FormatCurrency(RSImp!Valor_Juros, 2)
             End If
             
             If CDeb = 1 Then
                 .ParameterFields(15).AddCurrentValue "Desp. Cartão de Débito: " & FormatCurrency(RSImp!Pagar - RSImp!Custas - Distribuidor - Selo, 2)
             End If
             
             If RS!Editado = False Then
                 '.ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(RSProtesto!txFRJ, 2)
                 frjSoma = frjProt + frjAponta + frjCpd + frjIntima + (frjIntima * nAvalista) + frjDistrib + (frjEdital * nAvalEdital)
                 frcSoma = frcProt + frcAponta + frcCpd + frcIntima + (frcIntima * nAvalista) + frcDistrib + (frcEdital * nAvalEdital)
                 .ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(frjSoma, 2)
                 .ParameterFields(17).AddCurrentValue "Taxa FRC: " & FormatCurrency(frcSoma, 2)
             Else
                 frjSoma = frjProt + frjAponta + frjCpd + frjIntima + (frjIntima * nAvalista) + frjDistrib + frjEdital + (frjEdital * nAvalEdital)
                 frcSoma = frcProt + frcAponta + frcCpd + frcIntima + (frcIntima * nAvalista) + frcDistrib + frcEdital + (frcEdital * nAvalEdital)
                 .ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(frjSoma, 2)
                 .ParameterFields(17).AddCurrentValue "Taxa FRC: " & FormatCurrency(frcSoma, 2)
             End If
             DB.BeginTrans
             DB.Execute ("Update tblFinanceiro set FRJ='" & Replace(frjSoma, ",", ".") & "',FRC='" & Replace(frcSoma, ",", ".") & "' Where idFinanceiro='" & RSImp!idFinanceiro & "'")
             DB.CommitTrans
             SomaProt = eProt
             If RS!ContraProtesto = True Then
                 SomaProt = eProt + eCtProt
                 frjSoma = frjSoma + frjCtProt
                 frcSoma = frcSoma + frcCtProt
                 .ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(frjSoma, 2)
                 .ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(frcSoma, 2)
             End If
             .ParameterFields(18).AddCurrentValue "Selo Judicial: " & FormatCurrency(Selo, 2)

'                If IsNull(RSProtesto!txFRJ) Then
'                    .ParameterFields(19).AddCurrentValue "Distribuidor: " & FormatCurrency(vDistrib, 2)
'                Else
                 .ParameterFields(19).AddCurrentValue "Distribuidor: " & FormatCurrency(eDistrib, 2)
'                End If
             .ParameterFields(20).AddCurrentValue "Desp. Apontamento: " & FormatCurrency(eAponta, 2)
             .ParameterFields(21).AddCurrentValue "Desp. Intimação: " & FormatCurrency(eIntima + (eIntima * nAvalista), 2)
             
             .ParameterFields(22).AddCurrentValue "Desp. CPD: " & FormatCurrency(eCpd, 2)
                 
             If RS!Editado = True And nAvalEdital > 0 Then
                 .ParameterFields(23).AddCurrentValue "Edital: " & FormatCurrency(eEdital + (eEdital * nAvalEdital), 2)
             End If
             
             If RS!Editado = True And nAvalEdital = 0 Then
                 .ParameterFields(23).AddCurrentValue "Edital: " & FormatCurrency(eEdital, 2)
             End If
             
             'TT = FormatCurrency(RS!V_Protesto, 2)
             .ParameterFields(24).AddCurrentValue "Desp. Protesto: " & FormatCurrency(SomaProt, 2)
             .ParameterFields(25).AddCurrentValue "Adiantamento -: " & FormatCurrency(Adiantamento, 2)
             .ParameterFields(26).AddCurrentValue "TOTAL PAGO: " & FormatCurrency(vTotal, 2)
             .ParameterFields(27).AddCurrentValue "" & "Belém, " & Day(CDate(Format(Data_Ocorrencia, "00/00/0000"))) & " de " & MonthName(Month(CDate(Format(Data_Ocorrencia, "00/00/0000")))) & " de " & Year(CDate(Format(Data_Ocorrencia, "00/00/0000")))
             .ParameterFields(28).AddCurrentValue "Protocolo Dist.: " & RS!Protocolo_Dist
             '.ParameterFields(28).AddCurrentValue "Autenticacao: " & RSImp!Protocolo * (Day(Date) & Month(Date) & Year(Date))
             '.ParameterFields(29).AddCurrentValue "" & NumSelo1 & "-" & NumSelo2 & "-" & NumSelo3 & "-" & NumSelo4 & "-" & NumSelo5 & "-" & NumSelo6 & "-" & NumSelo7
         End With
         Set CRExportOptions = CrysRep.ExportOptions
         CRExportOptions.FormatType = crEFTPortableDocFormat
         CRExportOptions.DestinationType = crEDTDiskFile
         CRExportOptions.DiskFileName = FileLoca
         CrysRep.DisplayProgressDialog = False
         CrysRep.Export False
         Set CRExportOptions = Nothing
     End If
     RSProtesto.Close
    nResp = MsgBox("Imprimir o Recibo", vbQuestion + vbYesNo)
    If nResp = vbYes Then
        CrysRep.PrintOut False, 1
    End If
     Screen.MousePointer = 1
     'frmPrtCaixa.CRViewer1.ReportSource = CrysRep
     'frmPrtCaixa.CRViewer1.ViewReport
     'frmPrtCaixa.Show 1
     If Conn.State = 1 Then Conn.Close
     CaixaXML
'     EnvioXML
'<<< Apaga QRCode >>>
    RSPrt.Open "Select * From tblCaixa", DB, adOpenDynamic
        Do While Not RSPrt.EOF
            Kill RSPrt!QRCode
            RSPrt.MoveNext
        Loop
    RSPrt.Close
     
     DB.Execute "Delete From tblCaixa"
              
    End If
'<<< F I M  C U S T A S  D E  P R O T E S T O >>>

    DB.Execute ("UPDATE tblGuias SET Baixado='" & 1 & "'WHERE N_Guia='" & RSBaixa!N_Guia & "'")
    Else
        MsgBox "O Título " & RSBaixa!Protocolo & " não pode ser CANCELADO.", vbInformation
    End If
    If RS.State = 1 Then RS.Close
    If RSSeloDigital.State = 1 Then RSSeloDigital.Close
    If RSPostecipa.State = 1 Then RSPostecipa.Close
    If RSImp.State = 1 Then RSImp.Close
    If RSAdianta.State = 1 Then RSAdianta.Close
    RSBaixa.MoveNext
    Loop
    
'    MeuGrid
'    Carrega_Grid
'    MsgBox "Liquidação finalizada com sucesso!", vbInformation
        
    End If
    
    If Tipo_Ocorrencia = "RETIRADA" Then
    
        FRJ = 0
        FRC = 0
        totFRJ = 0
        totFRC = 0
        Multa = 0
        CPMF = 0
        Juros = 0
        SeloGeral = 0
        SeloCertidao = 0
        Mora = 0
        cxNomeRecibo = ""
        CancelaBanco = 0
    
        Do While Not RSBaixa.EOF
        RS.Open "Select * from tblTitulo where Protocolo_Cartorio=" & RSBaixa!Protocolo, DB, adOpenDynamic
    
        If RS.RecordCount = 0 Then
            MsgBox "O Título " & RSBaixa!Protocolo & " não pertence a este cartório.", vbInformation
            Libera = 1
            Exit Sub
        End If
    
        If RS!Tipo_Ocorrencia = "" Or RS!Tipo_Ocorrencia = "0" Or IsNull(RS!Tipo_Ocorrencia) Then
            Valor = RS!Saldo
            If Valor <= RSAponta!Faixa0 Then ValorCustas = FormatCurrency(RSAponta!Retirado0, 2): vRet = RSAponta!sPago0: atoPago = 171: nFaixa0 = 0
            If Valor <= RSAponta!Faixa1 And Valor > RSAponta!Faixa0 Then ValorCustas = FormatCurrency(RSAponta!Retirado1, 2): vRet = RSAponta!sPago1: atoPago = 172
            If Valor <= RSAponta!Faixa2 And Valor > RSAponta!Faixa1 Then ValorCustas = FormatCurrency(RSAponta!Retirado2, 2): vRet = RSAponta!sPago2: atoPago = 173
            If Valor <= RSAponta!Faixa3 And Valor > RSAponta!Faixa2 Then ValorCustas = FormatCurrency(RSAponta!Retirado3, 2): vRet = RSAponta!sPago3: atoPago = 174
            If Valor <= RSAponta!Faixa4 And Valor > RSAponta!Faixa3 Then ValorCustas = FormatCurrency(RSAponta!Retirado4, 2): vRet = RSAponta!sPago4: atoPago = 175
            If Valor <= RSAponta!Faixa5 And Valor > RSAponta!Faixa4 Then ValorCustas = FormatCurrency(RSAponta!Retirado5, 2): vRet = RSAponta!sPago5: atoPago = 176
            If Valor <= RSAponta!Faixa6 And Valor > RSAponta!Faixa5 Then ValorCustas = FormatCurrency(RSAponta!Retirado6, 2): vRet = RSAponta!sPago6: atoPago = 177
            If Valor > RSAponta!Faixa6 Then ValorCustas = FormatCurrency(RSAponta!Retirado7, 2): vRet = RSAponta!sPago7: atoPago = 178
                             
            eRet = vRet
            frcRet = Format(eRet * 0.025, "0.00")
            frjRet = Format(eRet * 0.15, "0.00")
            vRet = eRet + frcRet + frjRet
            
            eAponta = RSAponta!Apontamento
            frcAponta = Format(eAponta * 0.025, "0.00")
            frjAponta = Format(eAponta * 0.15, "0.00")
            vAponta = eAponta + frcAponta + frjAponta
            
            
            eIntima = RSAponta!Intimacao
            frcIntima = Format(eIntima * 0.025, "0.00")
            frjIntima = Format(eIntima * 0.15, "0.00")
            vIntima = eIntima + frcIntima + frjIntima
            
            eCanc = RSAponta!CancAponta
            frcCanc = Format(eCanc * 0.025, "0.00")
            frjCanc = Format(eCanc * 0.15, "0.00")
            vCanc = eCanc + frcCanc + frjCanc
            
            eDistrib = RSAponta!Distribuidor
            frcDistrib = Format(eDistrib * 0.025, "0.00")
            frjDistrib = Format(eDistrib * 0.15, "0.00")
            vDistrib = eDistrib + frcDistrib + frjDistrib
            
            eCpd = RSAponta!CPD
            frcCpd = Format(eCpd * 0.025, "0.00")
            frjCpd = Format(eCpd * 0.15, "0.00")
            vCpd = eCpd + frcCpd + frjCpd
            
            eEdital = RSAponta!V_Edital
            frcEdital = Format(eEdital * 0.025, "0.00")
            frjEdital = Format(eEdital * 0.15, "0.00")
            vEdital = eEdital + frcEdital + frjEdital
            
            eCtProt = RSAponta!ContraProtesto
            frcCtProt = Format(eCtProt * 0.025, "0.00")
            frjCtProt = Format(eCtProt * 0.15, "0.00")
            vCtProt = eCtProt + frcCtProt + frjCtProt
            
            Custas = ValorCustas
            Selo = RSAponta!Selo * 6
            Distribuidor = vDistrib
            Total = Custas + Selo + Distribuidor
            Juros = 0
            dOcorrencia = Date
            Adiantamento = FormatCurrency(0, 2)
            RSSeloDigital.Open "Select * From tblSeloDigital Where Usado='" & 0 & "' AND Tipo='" & 101 & "' Order By IdSelo", DB, adOpenDynamic
            
            If RSSeloDigital.RecordCount = 0 Then
                MsgBox "Sistema sem selos Digitais.", vbInformation, "Sem Selos"
                Screen.MousePointer = 1
                Exit Sub
            End If
            RSSeloDigital.MoveFirst
                
                
'<<< PREPARA A TABELA tblCaixa >>>
            totFRC = 0
            totFRJ = 0
    
'    <<< Busca Selo Digital >>>
            If RS!Especie_Tit = "CDA" Then
                RSBanco1.Open "Select * From tblBanco Where idBanco='" & RS!CodPortador & "'", DB, adOpenDynamic
                Portador = RSBanco1!Nome_Banco
                Cobranca = RSBanco1!Cobranca
                RSBanco1.Close
                If Cobranca = False Then
                    yato = 1
                    RSSeloDigital.Close
                    RSSeloDigital.Open "Select * From tblSeloDigital Where Usado='" & 0 & "' AND Tipo='" & 102 & "' Order By IdSelo", DB, adOpenDynamic
                    If RSSeloDigital.RecordCount = 0 Then
                        MsgBox "Sistema sem selos Digitais.", vbInformation, "Sem Selos"
                        Screen.MousePointer = 1
                        Exit Sub
                    End If
                    RSSeloDigital.MoveFirst
                Else
                    yato = 3
                    If RS!Editado = True Then
                        yato = yato + 1
                    End If
                    RSPostecipa.Open "Select * From tblSeloDigital Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Tipo='" & 113 & "' AND Postecipado='" & 0 & "'", DB, adOpenDynamic
                    If RSPostecipa.RecordCount > 0 Then
                        yato = 3
                        RSPostecipa.Close
                    End If
                End If
            Else
                yato = 3
                If RS!Editado = True Then
                    yato = yato + 1
                End If
                RSPostecipa.Open "Select * From tblSeloDigital Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Tipo='" & 113 & "' AND Postecipado='" & 0 & "'", DB, adOpenDynamic
                If RSPostecipa.RecordCount > 0 Then
                    yato = 3
                    RSPostecipa.Close
                End If
            End If
            For xAto = 1 To yato
            '<<< Arredonda os centavos >>>
            Select Case xAto
                Case 1
                    If RS!Especie_Tit = "CDA" And Cobranca = False Then
                        Emolumento = 0
                        Emol = 0
                        CodAto = atoPago
                        FRC = 0
                        FRJ = 0
                        Total = 0
                    Else
                        Emolumento = vRet
                        Emol = eRet
                        CodAto = atoPago
                        FRC = frcRet
                        FRJ = frjRet
                    End If
                Case 2
                    Emolumento = vCanc
                    Emol = eCanc
                    CodAto = 893
                    FRC = frcCanc
                    FRJ = frjCanc
                Case 3
                    Emolumento = vCpd
                    Emol = eCpd
                    CodAto = 967
                    FRC = frcCpd
                    FRJ = frjCpd
                Case 4
                    Emolumento = vEdital
                    Emol = eEdital
                    CodAto = 760
                    FRC = frcEdital
                    FRJ = frjEdital
            End Select
    
            FRJ = Format(FRJ, "0.00")
            FRC = Format(FRC, "0.00")
            totFRJ = totFRJ + FRJ
            totFRC = totFRC + FRC
    
            Codigo = RSSeloDigital!Codigo
            Serie = RSSeloDigital!Serie
            Tipo = RSSeloDigital!Tipo
            CodSeguranca = RSSeloDigital!CodSeguranca
                
            Value = "https://apps.tjpa.jus.br/ValidaSeloDigital/Detalhes/Resultado?codigoSelo=" & Codigo & "&serie=" & Serie & "&codigoTipoSelo=" & Tipo & "&codigoSeguranca=" & CodSeguranca & ""
            Caminho = ("\\" & Server & "\ProtestoSCP\QRCimg\Cx") & Tipo_Baixa & RS!Protocolo_Cartorio & "(" & xAto & ")" & ".bmp"
            GenerateBMP StrPtr(Caminho), StrPtr(Value), 3, 5, QualityLow
            
            RSCod.Open "Select * From tblEspecieTitulos WHERE Especie='" & RS!Especie_Tit & "'", DB, adOpenDynamic
            NCod = RSCod!Codigo
            RSCod.Close
            
            DB.BeginTrans
            DB.Execute ("Insert Into tblCaixa (Protocolo_Cartorio,CodAto,QRCode,N_Selo,Serie,Cod_Seg,Emolumentos,FRJ,FRC,NCod,TipoSelo) values ('" & RS!Protocolo_Cartorio & "','" & CodAto & "','" & Caminho & "','" & Codigo & "','" & Serie & "','" & CodSeguranca & "','" & Replace(Emol, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & NCod & "','" & Tipo & "')")
            DB.Execute ("Update tblSeloDigital set Usado='" & 1 & "',Protocolo='" & RS!Protocolo_Cartorio & "',Data_Uso='" & Format(Date, "mm/dd/yyyy") & "',Ato='" & CodAto & "' Where IdSelo='" & RSSeloDigital!IdSelo & "'")
            DB.CommitTrans
            RSSeloDigital.MoveNext
            Next
'<<< FIM PREPARA A TABELA tblCaixa >>>


'<<< ROTINA DE VERIFICAÇÃO DOS SELOS POSTECIPADOS >>>
            If RSPostecipa.State = 1 Then RSPostecipa.Close
            RSPostecipa.Open "Select * From tblSeloDigital Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Tipo='" & 113 & "' AND Postecipado='" & 0 & "'", DB, adOpenDynamic
            If RSPostecipa.RecordCount > 0 Then

                yato = RSPostecipa.RecordCount
        
                For xAto = 1 To yato
                '<<< Arredonda os centavos >>>
                Salva = 0
                Codigo = RSPostecipa!Codigo
                Serie = RSPostecipa!Serie
                Tipo = RSPostecipa!Tipo
                CodSeguranca = RSPostecipa!CodSeguranca
        
                'Apontamento
                If RSPostecipa!Ato = 152 Or RSPostecipa!Ato = 756 Then
                    Emolumento = vAponta
                    Emol = eAponta
                    CodAto = RSPostecipa!Ato
                    FRC = frcAponta
                    FRJ = frjAponta
                    Codigo = RSPostecipa!Codigo
                    Serie = RSPostecipa!Serie
                    Tipo = RSPostecipa!Tipo
                    CodSeguranca = RSPostecipa!CodSeguranca
                    Salva = 1
                End If
                'Distribuidor
                If RSPostecipa!Ato = 179 Or RSPostecipa!Ato = 755 Then
                    Emolumento = vDistrib
                    Emol = eDistrib
                    CodAto = RSPostecipa!Ato
                    FRC = frcDistrib
                    FRJ = frjDistrib
                    Codigo = RSPostecipa!Codigo
                    Serie = RSPostecipa!Serie
                    Tipo = RSPostecipa!Tipo
                    CodSeguranca = RSPostecipa!CodSeguranca
                    Salva = 1
                End If
                'Intimação
                If RSPostecipa!Ato = 162 Or RSPostecipa!Ato = 757 Then
                    Emolumento = vIntima
                    Emol = eIntima
                    CodAto = RSPostecipa!Ato
                    FRC = frcIntima
                    FRJ = frjIntima
                    Codigo = RSPostecipa!Codigo
                    Serie = RSPostecipa!Serie
                    Tipo = RSPostecipa!Tipo
                    CodSeguranca = RSPostecipa!CodSeguranca
                    Salva = 1
                End If
                'Edital
                If RSPostecipa!Ato = 164 Or RSPostecipa!Ato = 760 Then
                    Emolumento = vEdital
                    Emol = eEdital
                    CodAto = RSPostecipa!Ato
                    FRC = frcEdital
                    FRJ = frjEdital
                    Codigo = RSPostecipa!Codigo
                    Serie = RSPostecipa!Serie
                    Tipo = RSPostecipa!Tipo
                    CodSeguranca = RSPostecipa!CodSeguranca
                    Salva = 1
                End If
                
                FRJ = Format(FRJ, "0.00")
                FRC = Format(FRC, "0.00")
                totFRJ = totFRJ + FRJ
                totFRC = totFRC + FRC
        
        
                Value = "https://apps.tjpa.jus.br/ValidaSeloDigital/Detalhes/Resultado?codigoSelo=" & Codigo & "&serie=" & Serie & "&codigoTipoSelo=" & Tipo & "&codigoSeguranca=" & CodSeguranca & ""
                Caminho = ("\\" & Server & "\ProtestoSCP\QRCimg\Cx") & Tipo_Baixa & RS!Protocolo_Cartorio & "(" & xAto + 4 & ")" & ".bmp"
                GenerateBMP StrPtr(Caminho), StrPtr(Value), 3, 5, QualityLow
                
                RSCod.Open "Select * From tblEspecieTitulos WHERE Especie='" & RS!Especie_Tit & "'", DB, adOpenDynamic
                NCod = RSCod!Codigo
                RSCod.Close

                If Salva = 1 Then
                    DB.BeginTrans
                    DB.Execute ("Insert Into tblCaixa (Protocolo_Cartorio,CodAto,QRCode,N_Selo,Serie,Cod_Seg,Emolumentos,FRJ,FRC,NCod,TipoSelo) values ('" & RS!Protocolo_Cartorio & "','" & CodAto & "','" & Caminho & "','" & Codigo & "','" & Serie & "','" & CodSeguranca & "','" & Replace(Emol, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & NCod & "','" & Tipo & "')")
                    DB.Execute ("Update tblSeloDigital set Postecipado='" & 1 & "',Data_Uso='" & Format(Date, "mm/dd/yyyy") & "' Where IdSelo='" & RSPostecipa!IdSelo & "'")
                    DB.CommitTrans
                End If
                RSPostecipa.MoveNext
                Next
            End If
'<<< FIM DA ROTINA DE VERIFICAÇÃO DOS SELOS POSTECIPADOS >>>
         
            DB.Execute ("Insert Into tblFinanceiro (CodPortador,Data_Ocorrencia,Protocolo,Custas,Tipo_Ocorrencia,Pagar,Valor_Tit,Valor_Juros,Valor_CPMF,Valor_Selo,Valor_Mora,V_Multa,Valor_Distrib,Data_Retirada,Devedor,Usuario,Hora,QtdeSG,QtdeSC,Pagante,Especie_Tit) values ('" & RS!CodPortador & "','" & Format(Date, "ddmmyyyy") & "','" & RS!Protocolo_Cartorio & "','" & Replace(Custas, ",", ".") & "','" & 3 & "','" & Replace(Total, ",", ".") & "','" & Replace(Valor, ",", ".") & "','" & Replace(Juros, ",", ".") & "','" & Replace(CPMF, ",", ".") & "','" & Replace(Selo, ",", ".") & "','" & Replace(Mora, ",", ".") & "','" & Replace(Multa, ",", ".") & "','" & Replace(Distribuidor, ",", ".") & "','" & Format(Date, "mm/dd/yyyy") & "','" & RS!Devedor & "','" & User & "','" & Time & "','" & SeloGeral & "','" & SeloCertidao & "','" & cxNomeRecibo & "','" & RS!Especie_Tit & "')")
            DB.Execute ("Update tblTitulo set Baixado='" & 1 & "',Valor_Custas='" & Replace(Custas, ",", ".") & "',Valor_Selo='" & Replace(Selo, ",", ".") & "',Valor_Distrib='" & Replace(Distribuidor, ",", ".") & "',Tipo_Ocorrencia='" & 3 & "',Data_Ocorrencia='" & Format(Date, "ddmmyyyy") & "',Data_Retirada='" & Format(Date, "mm/dd/yyyy") & "',CancelaBanco='" & CancelaBanco & "' Where Protocolo_Dist='" & RSBaixa!Protocolo & "'")
            
            PrintConn
            
            RSImp.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "'AND Data_Ocorrencia='" & Format(Date, "ddmmyyyy") & "'AND Estorno Is Null", DB, adOpenDynamic
            If RSImp.RecordCount = 0 Then
                MsgBox "Sem dados para impressão!", vbInformation
                DB.RollbackTrans
                Screen.MousePointer = 1
                RSImp.Close
                Exit Sub
            End If
            If RS!CodPortador > 0 Then
                RSBanco1.Open "Select * From tblBanco Where idBanco='" & RS!CodPortador & "'", DB, adOpenDynamic
                Portador = RSBanco1!Nome_Banco
                RSBanco1.Close
            Else
                Portador = RS!Portador
            End If
        
            Set RSPrt = New ADODB.Recordset
            If RSPrt.State = 1 Then RSPrt.Close
            RSPrt.Open "Select * From tblCaixa", Conn
            
            Set CrysRep = CrysApp.OpenReport(App.Path & "\Recibo_RetiradaQRC.rpt")
            FileLoca = ("\\" & Server & "\ProtestoSCP\Caixa\" & Replace(Date, "/", "") & "\" & UCase(Trim(Replace(RS!Devedor, "/", ""))) & RS!Protocolo_Cartorio & "_RET_" & Right(Replace(Time, ":", ""), 2) & ".pdf")
            If Not RSPrt.EOF Then
                With CrysRep
                    Call .Database.Tables(1).SetDataSource(RSPrt)
                    .DiscardSavedData
                    .EnableParameterPrompting = False
                    .ReadRecords
                    .ParameterFields(1).AddCurrentValue "Recibo de RETIRADA " & FormatCurrency(RSImp!Pagar, 2)
                    .ParameterFields(2).AddCurrentValue "Protocolo: " & RSImp!Protocolo
                    .ParameterFields(3).AddCurrentValue "Recebemos de: " & Portador
                    Texto = SeparaExtenso(Extenso(RSImp!Pagar), 40)
                    .ParameterFields(4).AddCurrentValue "A importância de: " & Texto1 & Texto2 & " REFERENTES AS CUSTA DE RETIRADA SEM PROTESTO DO TÍTULO Nº " & RS!Num_Titulo & " NO VALOR DE " & FormatCurrency(RS!Saldo, 2) & " VENCIDO EM " & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Right(RS!Vencimento, 4) & " NOSSO NÚMERO " & RS!Nosso_Num
                    .ParameterFields(9).AddCurrentValue "Cedente " & RS!Cedente
                    .ParameterFields(10).AddCurrentValue "Contra " & RSImp!Devedor & " Conforme discriminacao abaixo."
                    If RSImp!Valor_Juros > 0 Then
                        .ParameterFields(15).AddCurrentValue "Juros: " & FormatCurrency(RSImp!Valor_Juros, 2)
                    End If
                    nData = CDate("01/12/2019")
                    If RS!Data_Apresenta > nData Then
                        .ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(totFRJ, 2)
                    Else
                        .ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(totFRJ + frjAponta + frjIntima + frjDistrib, 2)
                        totFRJ = totFRJ + frjAponta + frjIntima + frjDistrib
                    End If
                    If RS!Data_Apresenta > nData Then
                        .ParameterFields(17).AddCurrentValue "Taxa FRC: " & FormatCurrency(totFRC, 2)
                    Else
                        .ParameterFields(17).AddCurrentValue "Taxa FRC: " & FormatCurrency(totFRC + frcAponta + frcIntima + frcDistrib, 2)
                        totFRC = totFRC + frcAponta + frcIntima + frcDistrib
                    End If
                    DB.BeginTrans
                    DB.Execute ("Update tblFinanceiro set FRJ='" & Replace(totFRJ, ",", ".") & "',FRC='" & Replace(totFRC, ",", ".") & "' Where idFinanceiro='" & RSImp!idFinanceiro & "'")
                    DB.CommitTrans
                    If RSImp!Valor_Selo > 0 Then
                        .ParameterFields(18).AddCurrentValue "Selo Judicial: " & FormatCurrency(RSImp!Valor_Selo, 2)
                    End If
                    If RSImp!Valor_Distrib > 0 Then
                        .ParameterFields(19).AddCurrentValue "Distribuidor: " & FormatCurrency(eDistrib, 2)
                    End If
                    
                    .ParameterFields(20).AddCurrentValue "Desp. Apontamento: " & FormatCurrency(eAponta, 2)
                    .ParameterFields(21).AddCurrentValue "Desp. Intimação: " & FormatCurrency(eIntima + (eIntima * nAvalista), 2)
                    .ParameterFields(22).AddCurrentValue "Desp. Canc. Apont.: " & FormatCurrency(eCanc + eRet, 2)
                    .ParameterFields(23).AddCurrentValue "Desp. CPD: " & FormatCurrency(eCpd, 2)
                
                    If RS!Editado = True Then
                        .ParameterFields(24).AddCurrentValue "Edital: " & FormatCurrency(eEdital, 2)
                    End If
                    
                    '.ParameterFields(24).AddCurrentValue "TOTAL: " & FormatCurrency(RSImp!Pagar, 2)
                    .ParameterFields(25).AddCurrentValue "Adiantamento: " & FormatCurrency(Adiantamento, 2)

                    .ParameterFields(26).AddCurrentValue "TOTAL PAGO: " & FormatCurrency(RSImp!Pagar - Adiantamento, 2)
                    Data_Ocorrencia = Format(Date, "ddmmyyyy")
                    .ParameterFields(27).AddCurrentValue "" & "Belém, " & Day(CDate(Format(Data_Ocorrencia, "00/00/0000"))) & " de " & MonthName(Month(CDate(Format(Data_Ocorrencia, "00/00/0000")))) & " de " & Year(CDate(Format(Data_Ocorrencia, "00/00/0000")))
                    .ParameterFields(28).AddCurrentValue "Protocolo Dist.: " & RS!Protocolo_Dist
                End With
                Set CRExportOptions = CrysRep.ExportOptions
                CRExportOptions.FormatType = crEFTPortableDocFormat
                CRExportOptions.DestinationType = crEDTDiskFile
                CRExportOptions.DiskFileName = FileLoca
                CrysRep.DisplayProgressDialog = False
                CrysRep.Export False
                Set CRExportOptions = Nothing
            End If
            
            Screen.MousePointer = 1
'            frmPrtCaixa.CRViewer1.ReportSource = CrysRep
'            frmPrtCaixa.CRViewer1.ViewReport
'            frmPrtCaixa.Show 1
            If Conn.State = 1 Then Conn.Close
            CaixaXML
'            EnvioXML

    '<<< Apaga QRCode >>>
            RSPrt.Open "Select * From tblCaixa", DB, adOpenDynamic
                Do While Not RSPrt.EOF
                    Kill RSPrt!QRCode
                    RSPrt.MoveNext
                Loop
            RSPrt.Close
            
            DB.Execute "Delete From tblCaixa"
'    RSBaixa.MoveNext
'    Loop

'<<< F I M   R E T I R A D A >>>
                
    Else
        MsgBox "O Título " & RSBaixa!Protocolo & " não pode ser RETIRADO.", vbInformation
    End If
        DB.Execute ("UPDATE tblGuias SET Baixado='" & 1 & "'WHERE N_Guia='" & RSBaixa!N_Guia & "'")
        RSBaixa.MoveNext
        
        If RS.State = 1 Then RS.Close
        If RSSeloDigital.State = 1 Then RSSeloDigital.Close
        If RSPostecipa.State = 1 Then RSPostecipa.Close
        If RSImp.State = 1 Then RSImp.Close
        If RSAdianta.State = 1 Then RSAdianta.Close
        
        Loop
        
    End If
        
    MeuGrid
    Carrega_Grid
    MsgBox "Liquidação finalizada com sucesso!", vbInformation
    
    Exit Sub
Erro:
    If Err.Number = 53 Then
        Resume Next
    End If
    MsgBox "Erro de Sistema " & Err.Number & Err.Description, vbCritical, "Erro"
    Resume Next
    Close #1
'    DB.RollbackTrans
    Screen.MousePointer = 1

End Sub

Private Sub cmdImprimir_Click()
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
RS.Open "Select * From tblCalculo WHERE Usuario='" & User & "'", DB, adOpenDynamic
If RS.RecordCount = 0 Then
    RS.Close
    Exit Sub
End If

Dim CRExportOptions As Object
PrintConn
Set RSPrt = New ADODB.Recordset
If RSPrt.State = 1 Then RSPrt.Close
RSPrt.Open "Select * From tblCalculo WHERE Usuario='" & User & "'", Conn

Set CrysRep = CrysApp.OpenReport(App.Path & "\\CalculoCaixa.rpt")

FileLoca = ("\\" & Server & "\ProtestoSCP\Orcamentos\" & Trim(RSPrt!Devedor) & "_" & Replace(Date, "/", "") & ".pdf")

If Not RSPrt.EOF Then
    With CrysRep
        Call .Database.Tables(1).SetDataSource(RSPrt)
        .EnableParameterPrompting = False
        .DiscardSavedData
        .ReadRecords
    End With
Set CRExportOptions = CrysRep.ExportOptions
CRExportOptions.FormatType = crEFTPortableDocFormat
CRExportOptions.DestinationType = crEDTDiskFile
CRExportOptions.DiskFileName = FileLoca
CrysRep.DisplayProgressDialog = False
CrysRep.Export False
CrysRep.PrintOut False, 1
'frmPrtCaixa.CRViewer1.ReportSource = CrysRep
'frmPrtCaixa.CRViewer1.ViewReport
'frmPrtCaixa.Show 1

Set CRExportOptions = Nothing

End If
'DB.Execute "Delete from tblCalculo WHERE Usuario='" & User & "'"

End Sub

Private Sub cmdInserir_Click()
On Error GoTo Erro
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset

RS.Open "SELECT * FROM tblGuias ORDER BY id", DB, adOpenDynamic

If RS.RecordCount = 0 Then
    Cont = 1
    N_Guia = 1
Else
'    nResp = MsgBox("Nova Guia?", vbYesNo + vbQuestion)
    nResp = vbYes
    If nResp = vbYes Then
        RS.MoveLast
        Cont = RS!N_Guia
        Cont = Cont + 1
        N_Guia = Cont
    Else
        RS.MoveLast
        N_Guia = RS!N_Guia
    End If
End If

RS.Close
RS.Open "SELECT * FROM tblGuiaProvisoria WHERE USuario='" & User & "'ORDER BY id", DB, adOpenDynamic
If RS.RecordCount > 0 Then
    Do While Not RS.EOF
        DB.BeginTrans
        DB.Execute ("Insert Into tblGuias (Protocolo,Ocorrencia,Usuario,Devedor,Num_Devedor,N_Guia,Baixado,Tipo_Baixa,Pagar,Marca,DataEntrada) values ('" & RS!Protocolo & "','" & RS!Ocorrencia & "','" & User & "','" & Trim(RS!Devedor) & "','" & RS!Num_Devedor & "','" & N_Guia & "','" & 0 & "','" & RS!Tipo_Baixa & "','" & Replace(CDbl(RS!Pagar), ",", ".") & "','" & 0 & "','" & Format(Date, "mm/dd/yyyy") & "')")
        DB.CommitTrans
    RS.MoveNext
    Loop
    DB.Execute ("DELETE FROM tblCalculo WHERE Usuario='" & User & "'")
End If
    DB.Execute ("DELETE FROM tblGuiaProvisoria WHERE Usuario='" & User & "'")
    MeuGridTitulos
    Carrega_GridTitulos
    MeuGrid
    Carrega_Grid

Exit Sub
Erro:
    If Err.Number = -2147217873 Then
        MsgBox "O Protocolo n° " & RS!Protocolo & " já foi cadastrado em outra guia.", vbCritical
    Else
        MsgBox "Erro de sistema n° " & Err.Number & " - " & Err.Description, vbCritical
    End If
    DB.RollbackTrans
End Sub

Private Sub cmdInserirNaGuia_Click()
On Error GoTo Erro
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset

N_Guia = Trim(InputBox("Digite o Número da Guia para Incluir o Título."))
RS.Open "SELECT * FROM tblGuias WHERE N_Guia='" & N_Guia & "'AND Baixado=0 AND Marca=0 ORDER BY id", DB, adOpenDynamic

If RS.RecordCount = 0 Then
    MsgBox "N° de Guia " & N_Guia & " Inexistente ou Liquidada.", vbInformation
    RS.Close
    Exit Sub
End If

RS.Close
RS.Open "SELECT * FROM tblGuiaProvisoria WHERE USuario='" & User & "'ORDER BY id", DB, adOpenDynamic
If RS.RecordCount > 0 Then
    Do While Not RS.EOF
        DB.BeginTrans
        DB.Execute ("Insert Into tblGuias (Protocolo,Ocorrencia,Usuario,Devedor,Num_Devedor,N_Guia,Baixado,Tipo_Baixa,Pagar,Marca,DataEntrada) values ('" & RS!Protocolo & "','" & RS!Ocorrencia & "','" & User & "','" & Trim(RS!Devedor) & "','" & RS!Num_Devedor & "','" & N_Guia & "','" & 0 & "','" & RS!Tipo_Baixa & "','" & Replace(CDbl(RS!Pagar), ",", ".") & "','" & 0 & "','" & Format(Date, "mm/dd/yyyy") & "')")
        DB.CommitTrans
    RS.MoveNext
    Loop
    DB.Execute ("DELETE FROM tblCalculo WHERE Usuario='" & User & "'")
End If
    DB.Execute ("DELETE FROM tblGuiaProvisoria WHERE Usuario='" & User & "'")
    MeuGridTitulos
    Carrega_GridTitulos
    MeuGrid
    Carrega_Grid

Exit Sub
Erro:
    If Err.Number = -2147217873 Then
        MsgBox "O Protocolo n° " & RS!Protocolo & " já foi cadastrado em outra guia.", vbCritical
    Else
        MsgBox "Erro de sistema n° " & Err.Number & " - " & Err.Description, vbCritical
    End If
    DB.RollbackTrans

End Sub

Private Sub cmdPesquisar_Click()
    frmPesquisar.Show 1
    MeuGrid
    Carrega_Grid
End Sub

Private Sub cmdRecibo_Click()
    Dim sFilter As String
    sFilter = "Todos Arquivos (*.*)" & Chr(0) & "*.*" & Chr(0)
    Caminho = LocalizarArquivo("\\" & Server & "\ProtestoSCP\Caixa\" & Replace(Date, "/", ""), "Localizando Arquivo", sFilter)
    ShellExecute hWnd, "open", Caminho, vbNullString, vbNullString, conSwNo

End Sub

Private Sub Form_Activate()
MeuGridTitulos
Carrega_GridTitulos
MeuGrid
Carrega_Grid
End Sub

Private Sub Form_Load()
sDataInicial = Date
sDataFinal = Date
End Sub

Private Sub GridGuias_Click()
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim RSFinan As ADODB.Recordset
    Set RSFinan = New ADODB.Recordset
    Dim Cartao As String
    
    
    x = Me.GridGuias.RowSel 'pega a linha do grid clikada
    y = Me.GridGuias.ColSel 'pega a coluna do grid clicada
    
    If uCaixa = True Then
        Me.optDataLiquida.Visible = True
        If y <> 6 Then
            MsgBox "Clique no número da Guia.", vbInformation
        Else
            nResp = MsgBox("Confirma a Liquidação da Guia N° " & Me.GridGuias.TextMatrix(x, 6) & " no valor de " & Me.GridGuias.TextMatrix(x, 4), vbYesNo + vbQuestion)
            If nResp = vbYes Then
                Baixar
            End If
        End If
    Else
        If y = 7 Then
            If Me.GridGuias.TextMatrix(x, 7) = "Baixar" And Me.GridGuias.TextMatrix(x, 9) = "NÃO" Then
                Me.GridGuias.TextMatrix(x, 7) = ""
                Me.txtPagar = ""
                DB.Execute ("UPDATE tblGuias SET Marca='" & 0 & "'WHERE N_Guia='" & Me.GridGuias.TextMatrix(x, 6) & "'")
                Me.Refresh
            Else
                Total = 0
                RS.Open "SELECT * FROM tblGuias Where N_Guia='" & Me.GridGuias.TextMatrix(x, 6) & "'", DB, adOpenDynamic
                Do While Not RS.EOF
                    If RS!Marca = True Then
                        Me.GridGuias.TextMatrix(x, 7) = "Baixar"
                    Else
                        Me.GridGuias.TextMatrix(x, 7) = ""
                    End If
                    Total = RS!Pagar + Total
                RS.MoveNext
                Loop
                Me.txtPagar = FormatCurrency(Total, 2)
                DB.Execute ("UPDATE tblGuias SET Marca='" & 1 & "'WHERE N_Guia='" & Me.GridGuias.TextMatrix(x, 6) & "'")
                Me.Refresh
            End If
        End If
        
        If y = 6 Then
        MeuGridTitulos

            RS.Open "SELECT * FROM tblGuias WHERE N_Guia='" & Me.GridGuias.TextMatrix(x, 6) & "'ORDER BY N_Guia", DB, adOpenDynamic

            Me.txtTotalTitulos = RS.RecordCount
            Pagar = 0
            Do While Not RS.EOF

                        If RS!Marca = True Then
                            Marca = "Baixar"
                        Else
                            Marca = ""
                        End If
        
                        If RS!Baixado = True Then
                            Baixado = "SIM"
                        Else
                            Baixado = "NÃO"
                        End If
                        Pagar = Pagar + RS!Pagar
                
                RSFinan.Open "SELECT * FROM tblfinanceiro WHERE Protocolo = '" & RS!Protocolo & "'", DB, adOpenDynamic
                    If RSFinan!Cartao = True Then Cartao = "C" Else Cartao = "D"
                
                Me.GridTitulos.AddItem RS("Protocolo") & vbTab & RS("Ocorrencia") & vbTab & RS("Devedor") & vbTab & RS("Tipo_Baixa") & vbTab & FormatCurrency(RS("Pagar"), 2) & vbTab & RS("Usuario") & vbTab & RS("N_Guia") & vbTab & Marca & vbTab & RS("DataEntrada") & vbTab & Baixado & "  |  " & Cartao
                Me.Refresh
                RSFinan.Close
            RS.MoveNext
            Loop
            Me.txtPagar = FormatCurrency(Pagar, 2)

        End If
            
    End If
'    MeuGrid
'    Carrega_Grid
'    ZEBRAR
End Sub







Private Sub GridTitulos_Click()
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    
    x = Me.GridTitulos.RowSel 'pega a linha do grid clikada
    y = Me.GridTitulos.ColSel 'pega a coluna do grid clicada
    
    If uCaixa = True Then
        Me.optDataLiquida.Visible = True
        If y <> 6 Then
            MsgBox "Clique no número da Guia.", vbInformation
        Else
            nResp = MsgBox("Confirma a Liquidação da Guia N° " & Me.GridGuias.TextMatrix(x, 6) & " no valor de " & Me.GridGuias.TextMatrix(x, 4), vbYesNo + vbQuestion)
            If nResp = vbYes Then
                Baixar
            End If
        End If
    Else
        If y = 0 Then
            nResp = MsgBox("Confirma a exclusão do Protocolo " & Me.GridTitulos.TextMatrix(x, 0), vbYesNo + vbQuestion)
            If nResp = vbYes Then
                DB.Execute ("DELETE FROM tblGuiaProvisoria WHERE Protocolo='" & Me.GridTitulos.TextMatrix(x, 0) & "'")
                DB.Execute ("DELETE FROM tblCalculo WHERE Protocolo='" & Me.GridTitulos.TextMatrix(x, 0) & "'")
                DB.Execute ("DELETE FROM tblGuias WHERE Protocolo='" & Me.GridTitulos.TextMatrix(x, 0) & "'AND Baixado='" & 0 & "'")
                Me.MeuGridTitulos
                Me.Carrega_GridTitulos
            End If
        End If
    End If
    MeuGrid
    Carrega_Grid
    MeuGridTitulos
    Carrega_GridTitulos
End Sub



Private Sub txtProtocolo_GotFocus()
    Limpar
End Sub

Private Sub txtProtocolo_LostFocus()
On Error GoTo Erro

Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset

RS.Open "SELECT * FROM tblTitulo WHERE Protocolo_Cartorio ='" & Me.txtProtocolo & "'", DB, adOpenDynamic

If RS.RecordCount = 0 Then
    MsgBox "Protocolo inexistente.", vbInformation
    Exit Sub
End If

CDeb = 0
Me.txtDevedor = RS!Devedor
Me.txtNum_Devedor = RS!Num_Devedor
Tipo_Ocorrencia = RS!Tipo_Ocorrencia
If IsNull(RS!Protocolo_Dist) Then
    Me.txtProtocoloDist = 0
Else
    Me.txtProtocoloDist = RS!Protocolo_Dist
End If

Me.txtSaldo = RS!Saldo
Me.txtPagar = 0
If IsNull(Tipo_Ocorrencia) Then
    Tipo_Ocorrencia = ""
End If
Select Case Tipo_Ocorrencia
    Case Is = "": Me.txtOcorrencia = "SEM OCORRÊNCIA": Me.cboBaixa = "RETIRADA"
    Case Is = "A": Me.txtOcorrencia = "CANCELADO"
    Case Is = "1": Me.txtOcorrencia = "PAGO":
    Case Is = "2": Me.txtOcorrencia = "PROTESTADO":  Me.cboBaixa = "CANCELAMENTO"
    Case Is = "3": Me.txtOcorrencia = "RETIRADO"
    Case Is = "4": Me.txtOcorrencia = "SUSTADO"
    Case Is = "5": Me.txtOcorrencia = "DEVOLVIDO"
    Case Is = "6": Me.txtOcorrencia = "DEVOLVIDO"
End Select

If RS!Baixado = True Then
    MsgBox "Este título foi " & Me.txtOcorrencia, vbInformation
    Exit Sub
End If

Incluir

Exit Sub
Erro:
    MsgBox "Erro de Sistema. " & Err.Description & " - N° " & Err.Number, vbCritical

End Sub

Function Limpar()
'    Me.cmdIncluir.Enabled = False
'    Me.txtProtocolo = ""
    Me.txtDevedor = ""
    Me.txtOcorrencia = ""
    Me.txtNum_Devedor = ""
    Me.cboBaixa = ""
    Me.txtPagar = ""
    Me.txtProtocoloDist = ""
    Me.txtSaldo = ""
End Function

Public Sub MeuGrid()
    GridGuias.Clear
    GridGuias.FormatString = "Protocolo |Ocorrência |Devedor |Tipo Baixa |Total Pagar |Usuário |Nº Guia |Situação |Data |Liquidado "
    GridGuias.Cols = 10
    GridGuias.Rows = 1
    GridGuias.FixedCols = 0
    GridGuias.ColWidth(0) = 1000
    GridGuias.ColWidth(1) = 1800
    GridGuias.ColWidth(2) = 4000
    GridGuias.ColWidth(3) = 1500
    GridGuias.ColWidth(4) = 1100
    GridGuias.ColWidth(5) = 1300
    GridGuias.ColWidth(6) = 1000
    GridGuias.ColWidth(7) = 1100
    GridGuias.ColWidth(8) = 1000
    GridGuias.ColWidth(9) = 800
    GridGuias.ColAlignment(1) = 3
    GridGuias.ColAlignment(3) = 3
    GridGuias.ColAlignment(4) = 7
    GridGuias.CellAlignment = 1

End Sub

Public Sub Calcula_Baixa(RS As ADODB.Recordset)
On Error GoTo Erro
If Tipo_Baixa = "CANCELAMENTO" Then
'<<< CANCELAMENTO >>>
    Dim Aponta As ADODB.Recordset
    Set Aponta = New ADODB.Recordset
'    Dim RSUser As ADODB.Recordset
'    Set RSUser = New ADODB.Recordset
    Dim RSFinanceiro As ADODB.Recordset
    Set RSFinanceiro = New ADODB.Recordset
    Dim RSAvalista As ADODB.Recordset
    Set RSAvalista = New ADODB.Recordset
    Dim FRC As Double
    Dim FRJ As Double
    
    SeloGeral = 0
    SeloCertidao = 0
    
    Aponta.Open "Select * From tblApontamento", DB, adOpenDynamic
    
    eAponta = Aponta!Apontamento
    frcAponta = Format(eAponta * 0.025, "0.00")
    frjAponta = Format(eAponta * 0.15, "0.00")
    vAponta = eAponta + frcAponta + frjAponta
    issAponta = Format(eAponta * 0.05, "0.00")
    
    eIntima = Aponta!Intimacao
    frcIntima = Format(eIntima * 0.025, "0.00")
    frjIntima = Format(eIntima * 0.15, "0.00")
    vIntima = eIntima + frcIntima + frjIntima
    issIntima = Format(eIntima * 0.05, "0.00")
    
    eCanc = Aponta!CancAponta
    frcCanc = Format(eCanc * 0.025, "0.00")
    frjCanc = Format(eCanc * 0.15, "0.00")
    vCanc = eCanc + frcCanc + frjCanc
    issCanc = Format(eCanc * 0.05, "0.00")
    
    eDistrib = Aponta!Distribuidor
    frcDistrib = Format(eDistrib * 0.025, "0.00")
    frjDistrib = Format(eDistrib * 0.15, "0.00")
    vDistrib = eDistrib + frcDistrib + frjDistrib
    issDistrib = Format(eDistrib * 0.05, "0.00")
    
    eCpd = Aponta!CPD
    frcCpd = Format(eCpd * 0.025, "0.00")
    frjCpd = Format(eCpd * 0.15, "0.00")
    vCpd = eCpd + frcCpd + frjCpd
    issCpd = Format(eCpd * 0.05, "0.00")
    
    eEdital = Aponta!V_Edital
    frcEdital = Format(eEdital * 0.025, "0.00")
    frjEdital = Format(eEdital * 0.15, "0.00")
    vEdital = eEdital + frcEdital + frjEdital
    issEdital = Format(eEdital * 0.05, "0.00")
    
    eCtProt = Aponta!ContraProtesto
    frcCtProt = Format(eCtProt * 0.025, "0.00")
    frjCtProt = Format(eCtProt * 0.15, "0.00")
    vCtProt = eCtProt + frcCtProt + frjCtProt
    issCtProt = Format(eCtProt * 0.05, "0.00")
    
    Valor = RS!Saldo
    
    nFaixa0 = 0
    CalculoFaixas Aponta
'        If Valor <= Aponta!Faixa0 Then txtValorCustas = FormatCurrency(Aponta!Cancelado0, 2): vPago = Aponta!sPago0: atoCanc = 154
'        If Valor <= Aponta!Faixa1 And Valor > Aponta!Faixa0 Then txtValorCustas = FormatCurrency(Aponta!Cancelado1, 2): vPago = Aponta!sPago1: atoCanc = 155
'        If Valor <= Aponta!Faixa2 And Valor > Aponta!Faixa1 Then txtValorCustas = FormatCurrency(Aponta!Cancelado2, 2): vPago = Aponta!sPago2: atoCanc = 156
'        If Valor <= Aponta!Faixa3 And Valor > Aponta!Faixa2 Then txtValorCustas = FormatCurrency(Aponta!Cancelado3, 2): vPago = Aponta!sPago3: atoCanc = 157
'        If Valor <= Aponta!Faixa4 And Valor > Aponta!Faixa3 Then txtValorCustas = FormatCurrency(Aponta!Cancelado4, 2): vPago = Aponta!sPago4: atoCanc = 158
'        If Valor <= Aponta!Faixa5 And Valor > Aponta!Faixa4 Then txtValorCustas = FormatCurrency(Aponta!Cancelado5, 2): vPago = Aponta!sPago5: atoCanc = 159
'        If Valor <= Aponta!Faixa6 And Valor > Aponta!Faixa5 Then txtValorCustas = FormatCurrency(Aponta!Cancelado6, 2): vPago = Aponta!sPago6: atoCanc = 160
'        If Valor > Aponta!Faixa6 Then txtValorCustas = FormatCurrency(Aponta!Cancelado7, 2): vPago = Aponta!sPago7: atoCanc = 161
        
        ePago = vPago
        frcPago = Format(ePago * 0.025, "0.00")
        frjPago = Format(ePago * 0.15, "0.00")
        vPago = ePago + frcPago + frjPago
        issPago = Format(ePago * 0.05, "0.00")
        
        
        
        ISS = issPago + issCanc
    
        
        txtValorCustas = FormatCurrency(txtValorCustas - vCpd, 2)
        SeloGeral = 2
        nAtos = 2
        
        Custas = (txtValorCustas)
    
    cAto = atoCanc
    txtSeloJudicial = FormatCurrency(Aponta!Selo * 2, 2)
    Selo = (txtSeloJudicial)
    Total = Custas + Selo + ISS
    If CDeb = 1 Then
        If opSafra = True Then
            Taxa = (Aponta!txSafra / 100) + 1
        End If
        If opBradesco = True Then
            Taxa = (Aponta!txBradesco / 100) + 1
        End If
        If opItau = True Then
            Taxa = (Aponta!txItau / 100) + 1
        End If
'        Me.Label17.Visible = True
'        Me.Label17.Caption = "C. Débito"
'        frmCaixatxtValorCPMF.Visible = True
        txtValorCartorio = FormatCurrency((Custas + Selo) * (Taxa), 2)
        txtValorPagar = FormatCurrency(Total * (Taxa), 2)
        vTaxa = Format(Total * (Taxa - 1), "0.00")
        Total = Format(Total * (Taxa), "0.00")
        
    Else
'        Me.txtCancelamento = FormatCurrency((Custas + Selo), 2)
'        Me.Label7.Visible = True
'        Me.Label18.Visible = True
'        Me.Label20.Visible = True
'        Me.Line4.Visible = True
'        frmCaixatxtValorCPMF.Visible = True
        txtValorCartorio = FormatCurrency((Custas + Selo), 2)
        txtValorPagar = FormatCurrency(Total, 2)
        vTaxa = Format(Total, "0.00")
        Total = Format(Total, "0.00")
        txtTotal = FormatCurrency(Total, 2)
    End If
    Aponta.Close
    RS.Close

'<<< CUSTAS DE PROTESTO >>>
    vCheque = 0
    SeloGeral = 0
    SeloCertidao = 0
    nFaixa0 = 0
    nAvalIntima = 0
    nAvalista = 0
    nAvalEdital = 0
    
'    If Len(Me.txtProtocolo) > 6 Then
'        RS.Open "Select * From tblTitulo Where Protocolo_Dist='" & Me.txtProtocolo & "'", DB, adOpenDynamic
'        If RS.RecordCount = 0 Then
'        RS.Close
'        RS.Open "Select * From tblTitulo Where Protocolo_Cartorio='" & Me.txtProtocolo & "'", DB, adOpenDynamic
'            If RS.RecordCount = 0 Then
'                MsgBox "Título não encontrado.", vbInformation
'                RS.Close
'                Exit Sub
'            End If
'        End If
'    Else
        RS.Open "Select * From tblTitulo Where Protocolo_Cartorio='" & Me.txtProtocolo & "'", DB, adOpenDynamic
'    End If
    Me.txtDevedor = RS!Devedor
    Aponta.Open "Select * From tblApontamento", DB, adOpenDynamic

    If RS!Data_Apresenta < CDate("01/12/2019") And RS!Especie_Tit <> "CDA" Then
        For xAto = 1 To 2
        '<<< Arredonda os centavos >>>
        Select Case xAto
            Case 1
                Emolumento = vPago
                Emol = ePago
                FRC = frcPago
                FRJ = frjPago
'                Me.lblAto7 = "Canc.Prot."
                Ato7 = FormatCurrency(ePago, 2)
            Case 2
                Emolumento = vCanc
                Emol = eCanc
                FRC = frcCanc
                FRJ = frjCanc
'                Me.lblAto8 = "Canc.Aponta."
                Ato8 = FormatCurrency(eCanc, 2)
        End Select
        FRJ = Format(FRJ, "0.00")
        FRC = Format(FRC, "0.00")
        totFRJ = totFRJ + FRJ
        totFRC = totFRC + FRC
        Next

'        Exit Sub
    Else
    
'    RSFinanceiro.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "'And Tipo_Ocorrencia='" & 2 & "'And Estorno Is Null", DB, adOpenDynamic
'    If RSFinanceiro.RecordCount = 0 Then
'        MsgBox "O Título " & frmCaixatxtProtocolo & " não foi Protestado.", vbInformation
'        RSFinanceiro.Close
'        Aponta.Close
'        RS.Close
'        Exit Sub
'    End If
    Valor = RS!Saldo
    CalculoFaixaProtesto Aponta
'        If Valor <= Aponta!Faixa0 Then txtValorCustas = FormatCurrency(Aponta!Pago720, 2): vPago = Aponta!sPago0: atoPago = 171: vProt = Aponta!sProt0: atoCanc = 154: atoProt = 144: nFaixa0 = 0
'        If Valor <= Aponta!Faixa1 And Valor > Aponta!Faixa0 Then txtValorCustas = FormatCurrency(Aponta!Pago721, 2): vPago = Aponta!sPago1: atoPago = 172: vProt = Aponta!sProt1: atoCanc = 155: atoProt = 145
'        If Valor <= Aponta!Faixa2 And Valor > Aponta!Faixa1 Then txtValorCustas = FormatCurrency(Aponta!Pago722, 2): vPago = Aponta!sPago2: atoPago = 173: vProt = Aponta!sProt2: atoCanc = 156: atoProt = 146
'        If Valor <= Aponta!Faixa3 And Valor > Aponta!Faixa2 Then txtValorCustas = FormatCurrency(Aponta!Pago723, 2): vPago = Aponta!sPago3: atoPago = 174: vProt = Aponta!sProt3: atoCanc = 157: atoProt = 147
'        If Valor <= Aponta!Faixa4 And Valor > Aponta!Faixa3 Then txtValorCustas = FormatCurrency(Aponta!Pago724, 2): vPago = Aponta!sPago4: atoPago = 175: vProt = Aponta!sProt4: atoCanc = 158: atoProt = 148
'        If Valor <= Aponta!Faixa5 And Valor > Aponta!Faixa4 Then txtValorCustas = FormatCurrency(Aponta!Pago725, 2): vPago = Aponta!sPago5: atoPago = 176: vProt = Aponta!sProt5: atoCanc = 159: atoProt = 149
'        If Valor <= Aponta!Faixa6 And Valor > Aponta!Faixa5 Then txtValorCustas = FormatCurrency(Aponta!Pago726, 2): vPago = Aponta!sPago6: atoPago = 177: vProt = Aponta!sProt6: atoCanc = 160: atoProt = 150
'        If Valor > Aponta!Faixa6 Then txtValorCustas = FormatCurrency(Aponta!Pago727, 2): vPago = Aponta!sPago7: atoPago = 178: vProt = Aponta!sProt7: atoCanc = 161: atoProt = 151

        ePago = vPago
        frcPago = Format(ePago * 0.025, "0.00")
        frjPago = Format(ePago * 0.15, "0.00")
        vPago = ePago + frcPago + frjPago
                
        eProt = vProt
        Dim frjProt As Currency
        frcProt = Format(eProt * 0.025, "0.00")
        frjProt = Format(eProt * 0.15, "0.00")
        vProt = eProt + frcProt + frjProt
        issProt = Format(eProt * 0.05, "0.00")

'        SeloGeral = RSFinanceiro!QtdeSG
    
    If RS!Editado = True Then
        nAtos = 5
        txtValorCustas = FormatCurrency((vProt + vAponta + vCpd + vIntima + vEdital) + Custas, 2)
    Else
        nAtos = 4
        txtValorCustas = FormatCurrency((vProt + vAponta + vCpd + vIntima) + Custas, 2)
    End If
    
    If RS!ContraProtesto = True And RS!Editado = True Then
        nAtos = 6
        ContraProtesto = 1
        txtValorCustas = FormatCurrency((vProt + vAponta + vCpd + vIntima + vEdital + vCtProt) + Custas, 2)
    End If
    
    If RS!ContraProtesto = True And RS!Editado = False Then
        nAtos = 5
        ContraProtesto = 1
        txtValorCustas = FormatCurrency((vProt + vAponta + vCpd + vIntima + vCtProt) + Custas, 2)
    End If
    
    'Verifica se tem Avalista
    RSAvalista.Open "Select * From tblAvalista Where Protocolo_Cartorio='" & RS!Protocolo_Cartorio & "'", DB, adOpenDynamic
    
    Do While Not RSAvalista.EOF
        If RSAvalista!Editado = True Then
            nAvalEdital = nAvalEdital + 1
            txtValorCustas = txtValorCustas + vEdital
            nAtos = nAtos + 1
        End If
        If RSAvalista!Intimado = True Then
            nAvalIntima = nAvalIntima + 1
            txtValorCustas = txtValorCustas + vIntima
            nAtos = nAtos + 1
        End If
        nAvalista = nAvalista + 1
        RSAvalista.MoveNext
    Loop
    
    RSAvalista.Close
    
    
    Valor = RS!Saldo
    txtSeloJudicial = FormatCurrency((((nAtos + 1) * Aponta!Selo)) + Selo, 2)
    Distribuidor = FormatCurrency(vDistrib, 2)
    txtValorDistribuidor = FormatCurrency(Distribuidor, 2)
    Custas = (txtValorCustas)
    Selo = (txtSeloJudicial)
    Total = Custas + Selo + (Distribuidor)
    
    End If
    If CDeb = 1 Then
        If opSafra = True Then
            Taxa = (Aponta!txSafra / 100) + 1
        End If
        If opBradesco = True Then
            Taxa = (Aponta!txBradesco / 100) + 1
        End If
        If opItau = True Then
            Taxa = (Aponta!txItau / 100) + 1
        End If
        frmCaixaLabel17.Visible = True
        frmCaixaLabel17.Caption = "C. Débito"
        frmCaixatxtValorCPMF.Visible = True
        frmCaixatxtValorCartorio = FormatCurrency((Custas + Selo + Distribuidor) * Taxa, 2)
        frmCaixatxtValorPagar = FormatCurrency(Total * Taxa, 2)
        vTaxa = Format(Total * (Taxa - 1), "0.00")
        Total = Format(Total * Taxa, "0.00")
        frmCaixatxtValorCPMF = FormatCurrency(vTaxa, 2)
    Else
        txtValorCartorio = FormatCurrency((Custas + Selo), 2)
        txtValorPagar = FormatCurrency(Total, 2)
        vTaxa = Format(Total, "0.00")
        Total = Format(Total, "0.00")
        txtTotal = FormatCurrency(Total, 2)
'        Me.txtCustasProtesto = FormatCurrency((Total), 2)
'        Me.txtTotal = FormatCurrency(Total + Me.txtCancelamento, 2)
    End If
'    If RS!Data_Apresenta > CDate("01/12/2019") Then
        yato = 1
        If RS!ContraProtesto = True Then
            yato = 2
            ContraProtesto = 1
        End If
        
        For xAto = 1 To yato
        '<<< Arredonda os centavos >>>
                
        Select Case xAto
            Case 1
                Emolumento = vCpd
                Emol = eCpd
                FRC = frcCpd
                FRJ = frjCpd
                'lblAto1 = "CPD"
                Ato1 = FormatCurrency(eCpd, 2)
            Case 2
                Emolumento = vCtProt + vCpd
                Emol = eCtProt + eCpd
                FRC = frcCtProt
                FRJ = frjCtProt
'                Me.lblAto1 = "CPD/Cont.Prot."
                Ato1 = FormatCurrency(Emol, 2)
        End Select
        
        FRJ = Format(FRJ, "0.00")
        FRC = Format(FRC, "0.00")
        totFRJ = totFRJ + FRJ
        totFRC = totFRC + FRC
        Next

        If RS!Editado = True Then
            yato = 5
        Else
            yato = 4
        End If

        For xAto = 1 To yato
        'Protesto
        If xAto = 1 Then
            Emolumento = vProt
            Emol = eProt
            FRC = frcProt
            FRJ = frjProt
'            Me.lblAto2 = "Protesto"
            Ato2 = FormatCurrency(eProt, 2)
            Salva = 1
        End If
        'Apontamento
        If xAto = 2 Then
            Emolumento = vAponta
            Emol = eAponta
            FRC = frcAponta
            FRJ = frjAponta
'            Me.lblAto3 = "Apontamento"
            Ato3 = FormatCurrency(eAponta, 2)
            Salva = 1
        End If
        'Distribuidor
        If xAto = 3 Then
            Emolumento = vDistrib
            Emol = eDistrib
            FRC = frcDistrib
            FRJ = frjDistrib
'            Me.lblAto4 = "Distribuidor"
            Ato4 = FormatCurrency(eDistrib, 2)
            Salva = 1
        End If
        'Intimação
        If xAto = 4 Then
            Emolumento = vIntima
            Emol = eIntima
'            Me.lblAto5 = "Intimação"
            If nAvalIntima > 0 Then
                FRC = frcIntima * (nAvalIntima + 1)
                FRJ = frjIntima * (nAvalIntima + 1)
                Ato5 = FormatCurrency(eIntima * (nAvalIntima + 1), 2)
            Else
                FRC = frcIntima
                FRJ = frjIntima
                Ato5 = FormatCurrency(eIntima, 2)
            End If
            Salva = 1
        End If
        'Edital
        If xAto = 5 Then
            Emolumento = vEdital
            Emol = eEdital
'            Me.lblAto6 = "Edital"
            If nAvalEdital > 0 Then
                FRC = frcEdital * (nAvalEdital + 1)
                FRJ = frjEdital * (nAvalEdital + 1)
                Ato6 = FormatCurrency(eEdital * (nAvalEdital + 1), 2)
            Else
                FRC = frcEdital
                FRJ = frjEdital
                Ato6 = FormatCurrency(eEdital, 2)
            End If
            Salva = 1
        Else
'            Me.lblAto6 = "Edital"
            Ato6 = FormatCurrency(0, 2)
        End If
        
        FRJ = Format(FRJ, "0.00")
        FRC = Format(FRC, "0.00")
        totFRJ = totFRJ + FRJ
        totFRC = totFRC + FRC
        Next

        If RS!Editado = True Then
            totISS = issProt + issAponta + issCpd + issIntima + issDistrib + ISS + issEdital
        Else
            totISS = issProt + issAponta + issCpd + issIntima + issDistrib + ISS
        End If
        
        For xAto = 1 To 2
        '<<< Arredonda os centavos >>>
        Select Case xAto
            Case 1
                Emolumento = vPago
                Emol = ePago
                FRC = frcPago
                FRJ = frjPago
'                Me.lblAto7 = "Canc.Prot."
                Ato7 = FormatCurrency(ePago, 2)
            Case 2
                Emolumento = vCanc
                Emol = eCanc
                FRC = frcCanc
                FRJ = frjCanc
'                Me.lblAto8 = "Canc.Aponta."
                Ato8 = FormatCurrency(eCanc, 2)
        End Select
        FRJ = Format(FRJ, "0.00")
        FRC = Format(FRC, "0.00")
        totFRJ = totFRJ + FRJ
        totFRC = totFRC + FRC
        Next
        Ato9 = totFRJ
        Ato10 = totFRC
        
    If RS!Data_Apresenta < CDate("01/12/2019") And RS!Especie_Tit <> "CDA" Then
        totISS = 0
    End If
    txtTotal = txtTotal + totISS
    Me.txtPagar = txtValorPagar + totISS
'    End If
'    nResp = MsgBox("Colocar na Memória para Impressão", vbYesNo + vbQuestion)
'    If nResp = vbYes Then
        DB.Execute ("Insert Into tblGuiaProvisoria (N_Guia,Protocolo,DataEntrada,Ocorrencia,Pagar,Devedor,Custas,Selo,Valor_Titulo,Ato1,Ato2,Ato3,Ato4,Ato5,Ato6,Ato7,Ato8,Ato9,Ato10,Usuario,Num_Devedor,Tipo_Baixa) values ('" & 0 & "','" & txtProtocolo & "','" & Format(Date, "mm/dd/yyyy") & "','" & Me.txtOcorrencia & "','" & Replace(CDbl(txtTotal), ",", ".") & "','" & Trim(txtDevedor) & "','" & Replace(CDbl(Custas), ",", ".") & "','" & Replace(CDbl(Selo), ",", ".") & "','" & Replace(CDbl(Valor), ",", ".") & "','" & Replace(CDbl(Ato1), ",", ".") & "','" & Replace(CDbl(Ato2), ",", ".") & "','" & Replace(CDbl(Ato3), ",", ".") & "','" & Replace(CDbl(Ato4), ",", ".") & "','" & Replace(CDbl(Ato5), ",", ".") & "','" & Replace(CDbl(Ato6), ",", ".") & "','" & Replace(CDbl(Ato7), ",", ".") & "','" & Replace(CDbl(Ato8), ",", ".") & "','" & Replace(CDbl(totFRJ), ",", ".") & "','" & Replace(CDbl(totFRC), ",", ".") & "','" & User & "','" & RS!Num_Devedor & "','" & Tipo_Baixa & "')")
        DB.Execute ("Insert Into tblCalculo (Protocolo,Tipo_Ocorrencia,Valor,Devedor,Custas,Selo,Valor_Titulo,Ato1,Ato2,Ato3,Ato4,Ato5,Ato6,Ato7,Ato8,Ato9,Ato10,Usuario,ISS) values ('" & txtProtocolo & "','" & Tipo_Baixa & "','" & Replace(CDbl(txtTotal), ",", ".") & "','" & Trim(txtDevedor) & "','" & Replace(CDbl(Custas), ",", ".") & "','" & Replace(CDbl(Selo), ",", ".") & "','" & Replace(CDbl(Valor), ",", ".") & "','" & Replace(CDbl(Ato1), ",", ".") & "','" & Replace(CDbl(Ato2), ",", ".") & "','" & Replace(CDbl(Ato3), ",", ".") & "','" & Replace(CDbl(Ato4), ",", ".") & "','" & Replace(CDbl(Ato5), ",", ".") & "','" & Replace(CDbl(Ato6), ",", ".") & "','" & Replace(CDbl(Ato7), ",", ".") & "','" & Replace(CDbl(Ato8), ",", ".") & "','" & Replace(CDbl(Ato9), ",", ".") & "','" & Replace(CDbl(Ato10), ",", ".") & "','" & User & "','" & Replace(CDbl(totISS), ",", ".") & "')")
'    End If
End If

If Tipo_Baixa = "RETIRADA" Then


'<<< RETIRADA >>>
    Dim RSAval As ADODB.Recordset
    Set RSAval = New ADODB.Recordset
    Dim RSBanco As ADODB.Recordset
    Set RSBanco = New ADODB.Recordset
    Dim RSAdianta As ADODB.Recordset
    Set RSAdianta = New ADODB.Recordset
    Dim nDataProt As Date
    Dim RSAponta As ADODB.Recordset
    Set RSAponta = New ADODB.Recordset
    Dim RSSeloDigital As ADODB.Recordset
    Set RSSeloDigital = New ADODB.Recordset
    Dim RSSeloGratuito As ADODB.Recordset
    Set RSSeloGratuito = New ADODB.Recordset
    Dim RSProtesto As ADODB.Recordset
    Set RSProtesto = New ADODB.Recordset
    Dim RSPostecipa As ADODB.Recordset
    Set RSPostecipa = New ADODB.Recordset
    Dim RSCod As ADODB.Recordset
    Set RSCod = New ADODB.Recordset
    Dim dOcorrencia As Date
    Dim RSImp As ADODB.Recordset
    Set RSImp = New ADODB.Recordset
    Dim RSBanco1 As ADODB.Recordset
    Set RSBanco1 = New ADODB.Recordset
    Dim Portador As String
    
    FRJ = 0
    FRC = 0
    totFRJ = 0
    totFRC = 0
    Multa = 0
    CPMF = 0
    Juros = 0
    SeloGeral = 0
    SeloCertidao = 0
    Mora = 0
    cxNomeRecibo = ""
    CancelaBanco = 0
    
'<<< Abre Tabela de Apontamento >>>
    RSAponta.Open "Select * From tblApontamento", DB, adOpenDynamic

'<<< Busca o Protocolo na Tabela tblTitulo >>>
'    RS.Open "Select * From tblTitulo Where Protocolo_Cartorio='" & RSBaixa!Protocolo & "'AND Tipo_Ocorrencia<>'" & 2 & "'AND Tipo_Ocorrencia<>'" & 1 & "'", DB, adOpenDynamic
'    If Len(RSBaixa!Protocolo) > 6 Then
'        RS.Open "Select * From tblTitulo Where Protocolo_Dist='" & RSBaixa!Protocolo & "'", DB, adOpenDynamic
'    Else
'        RS.Open "Select * From tblTitulo Where Protocolo_Cartorio='" & RSBaixa!Protocolo & "'", DB, adOpenDynamic
'    End If
'    If Len(RS!Protocolo_Cartorio) > 6 Then
'        RS.Open "Select * from tblTitulo where Protocolo_Dist=" & RS!Protocolo_Cartorio, DB, adOpenDynamic
'        If RS.RecordCount = 0 Then
'            RS.Close
'            RS.Open "Select * from tblTitulo where Protocolo_Cartorio=" & RS!Protocolo_Cartorio, DB, adOpenDynamic
'        Else
'            If RSBaixa!Protocolo < RS!Protocolo_Cartorio Then
'                RS.Close
'                RS.Open "Select * from tblTitulo where Protocolo_Cartorio=" & RS!Protocolo_Cartorio, DB, adOpenDynamic
'            End If
'        End If
'    Else
'        RS.Open "Select * from tblTitulo where Protocolo_Cartorio=" & RSBaixa!Protocolo, DB, adOpenDynamic
'    End If

    If RS.RecordCount = 0 Then
        MsgBox "O Título " & RSBaixa!Protocolo & " não pertence a este cartório.", vbInformation
        Libera = 1
        Exit Sub
    Else
    If RS!Tipo_Ocorrencia = "" Or RS!Tipo_Ocorrencia = "0" Or IsNull(RS!Tipo_Ocorrencia) Then
        Valor = RS!Saldo
        Tipo_Baixa = "Retirada"
        CalculoFaixas RSAponta
'        If Valor <= RSAponta!Faixa0 Then ValorCustas = FormatCurrency(RSAponta!Retirado0, 2): vRet = RSAponta!sPago0: atoPago = 171: nFaixa0 = 0
'        If Valor <= RSAponta!Faixa1 And Valor > RSAponta!Faixa0 Then ValorCustas = FormatCurrency(RSAponta!Retirado1, 2): vRet = RSAponta!sPago1: atoPago = 172
'        If Valor <= RSAponta!Faixa2 And Valor > RSAponta!Faixa1 Then ValorCustas = FormatCurrency(RSAponta!Retirado2, 2): vRet = RSAponta!sPago2: atoPago = 173
'        If Valor <= RSAponta!Faixa3 And Valor > RSAponta!Faixa2 Then ValorCustas = FormatCurrency(RSAponta!Retirado3, 2): vRet = RSAponta!sPago3: atoPago = 174
'        If Valor <= RSAponta!Faixa4 And Valor > RSAponta!Faixa3 Then ValorCustas = FormatCurrency(RSAponta!Retirado4, 2): vRet = RSAponta!sPago4: atoPago = 175
'        If Valor <= RSAponta!Faixa5 And Valor > RSAponta!Faixa4 Then ValorCustas = FormatCurrency(RSAponta!Retirado5, 2): vRet = RSAponta!sPago5: atoPago = 176
'        If Valor <= RSAponta!Faixa6 And Valor > RSAponta!Faixa5 Then ValorCustas = FormatCurrency(RSAponta!Retirado6, 2): vRet = RSAponta!sPago6: atoPago = 177
'        If Valor > RSAponta!Faixa6 Then ValorCustas = FormatCurrency(RSAponta!Retirado7, 2): vRet = RSAponta!sPago7: atoPago = 178
                         
        ValorCustas = txtValorCustas
                         
        eRet = vRet
        frcRet = Format(eRet * 0.025, "0.00")
        frjRet = Format(eRet * 0.15, "0.00")
        vRet = eRet + frcRet + frjRet
        issRet = Format(eRet * 0.05, "0.00")
               
        eAponta = RSAponta!Apontamento
        frcAponta = Format(eAponta * 0.025, "0.00")
        frjAponta = Format(eAponta * 0.15, "0.00")
        vAponta = eAponta + frcAponta + frjAponta
        issAponta = Format(eAponta * 0.05, "0.00")
        
        eIntima = RSAponta!Intimacao
        frcIntima = Format(eIntima * 0.025, "0.00")
        frjIntima = Format(eIntima * 0.15, "0.00")
        vIntima = eIntima + frcIntima + frjIntima
        issIntima = Format(eIntima * 0.05, "0.00")
        
        eCanc = RSAponta!CancAponta
        frcCanc = Format(eCanc * 0.025, "0.00")
        frjCanc = Format(eCanc * 0.15, "0.00")
        vCanc = eCanc + frcCanc + frjCanc
        issCanc = Format(eCanc * 0.05, "0.00")
        
        eDistrib = RSAponta!Distribuidor
        frcDistrib = Format(eDistrib * 0.025, "0.00")
        frjDistrib = Format(eDistrib * 0.15, "0.00")
        vDistrib = eDistrib + frcDistrib + frjDistrib
        issDistrib = Format(eDistrib * 0.05, "0.00")
        
        eCpd = RSAponta!CPD
        frcCpd = Format(eCpd * 0.025, "0.00")
        frjCpd = Format(eCpd * 0.15, "0.00")
        vCpd = eCpd + frcCpd + frjCpd
        issCpd = Format(eCpd * 0.05, "0.00")
        
        eEdital = RSAponta!V_Edital
        frcEdital = Format(eEdital * 0.025, "0.00")
        frjEdital = Format(eEdital * 0.15, "0.00")
        vEdital = eEdital + frcEdital + frjEdital
        issEdital = Format(eEdital * 0.05, "0.00")
        
        eCtProt = RSAponta!ContraProtesto
        frcCtProt = Format(eCtProt * 0.025, "0.00")
        frjCtProt = Format(eCtProt * 0.15, "0.00")
        vCtProt = eCtProt + frcCtProt + frjCtProt
        issCtProt = Format(eCtProt * 0.05, "0.00")
                                                            
        If RS!Editado = False Or IsNull(RS!Editado) Then
            issEdital = 0
            vEdital = 0
            Selo = RSAponta!Selo * 6
        Else
            Selo = RSAponta!Selo * 7
        End If
        ISS = issRet + issIntima + issDistrib + issAponta + issCpd + issCanc + issEdital
                    
        Custas = ValorCustas
'        Selo = RSAponta!Selo * 6
        Distribuidor = vDistrib
        Total = Custas + Selo + Distribuidor + ISS + vEdital
        Juros = 0
        dOcorrencia = Date
        Adiantamento = FormatCurrency(0, 2)
        RSSeloDigital.Open "Select * From tblSeloDigital Where Usado='" & 0 & "' AND Tipo='" & 101 & "' Order By IdSelo", DB, adOpenDynamic

        If RSSeloDigital.RecordCount = 0 Then
            MsgBox "Sistema sem selos Digitais.", vbInformation, "Sem Selos"
            Screen.MousePointer = 1
            Exit Sub
        End If
        RSSeloDigital.MoveFirst
                
                
'<<< PREPARA A TABELA tblCaixa >>>
    totFRC = 0
    totFRJ = 0
    
'    <<< Busca Selo Digital >>>
    If RS!Especie_Tit = "CDA" Then
        RSBanco1.Open "Select * From tblBanco Where idBanco='" & RS!CodPortador & "'", DB, adOpenDynamic
        Portador = RSBanco1!Nome_Banco
        Cobranca = RSBanco1!Cobranca
        RSBanco1.Close
        If Cobranca = False Then
            yato = 1
            RSSeloDigital.Close
            RSSeloDigital.Open "Select * From tblSeloDigital Where Usado='" & 0 & "' AND Tipo='" & 102 & "' Order By IdSelo", DB, adOpenDynamic
            If RSSeloDigital.RecordCount = 0 Then
                MsgBox "Sistema sem selos Digitais.", vbInformation, "Sem Selos"
                Screen.MousePointer = 1
                Exit Sub
            End If
            RSSeloDigital.MoveFirst
        Else
            yato = 3
            If RS!Editado = True Then
                yato = yato + 1
            End If
            RSPostecipa.Open "Select * From tblSeloDigital Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Tipo='" & 113 & "' AND Postecipado='" & 0 & "'", DB, adOpenDynamic
            If RSPostecipa.RecordCount > 0 Then
'                yato = 3
                RSPostecipa.Close
            End If
        End If
    Else
        yato = 3
        If RS!Editado = True Then
            yato = yato + 1
        End If
    RSPostecipa.Open "Select * From tblSeloDigital Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Tipo='" & 113 & "' AND Postecipado='" & 0 & "'", DB, adOpenDynamic
    If RSPostecipa.RecordCount > 0 Then
'        yato = 3
        RSPostecipa.Close
    End If
    End If
        For xAto = 1 To yato
        '<<< Arredonda os centavos >>>
        Select Case xAto
            Case 1
                If RS!Especie_Tit = "CDA" And Cobranca = False Then
                    Emolumento = 0
                    Emol = 0
                    CodAto = atoPago
                    FRC = 0
                    FRJ = 0
                    Total = 0
                    Ato7 = 0
                Else
                    Emolumento = vRet
                    Emol = eRet
                    CodAto = atoPago
                    FRC = frcRet
                    FRJ = frjRet
                    Ato7 = FormatCurrency(eRet, 2)
                End If
            Case 2
                Emolumento = vCanc
                Emol = eCanc
                CodAto = 893
                FRC = frcCanc
                FRJ = frjCanc
                Ato8 = FormatCurrency(eCanc, 2)
            Case 3
'                Emolumento = vCpd
'                Emol = eCpd
'                CodAto = 180
                FRC = 0
                FRJ = 0
'                Ato1 = FormatCurrency(eCpd, 2)
            Case 4
                Emolumento = vEdital
                Emol = eEdital
                CodAto = 760
                FRC = frcEdital
                FRJ = frjEdital
                Ato6 = FormatCurrency(eEdital, 2)
        End Select

'        FRJ = Format(Emolumento, "0.00") * 0.15
'        FRC = Format(Emolumento, "0.00") * 0.025
        FRJ = Format(FRJ, "0.00")
        FRC = Format(FRC, "0.00")
        totFRJ = totFRJ + FRJ
        totFRC = totFRC + FRC
'
'        Codigo = RSSeloDigital!Codigo
'        Serie = RSSeloDigital!Serie
'        Tipo = RSSeloDigital!Tipo
'        CodSeguranca = RSSeloDigital!CodSeguranca
'
'        Value = "https://apps.tjpa.jus.br/ValidaSeloDigital/Detalhes/Resultado?codigoSelo=" & Codigo & "&serie=" & Serie & "&codigoTipoSelo=" & Tipo & "&codigoSeguranca=" & CodSeguranca & ""
'        Caminho = ("\\" & Server & "\QRCimg\Cx") & Tipo_Baixa & RS!Protocolo_Cartorio & "(" & xAto & ")" & ".bmp"
'        GenerateBMP StrPtr(Caminho), StrPtr(Value), 3, 5, QualityLow
        
'        RSCod.Open "Select * From tblEspecieTitulos WHERE Especie='" & RS!Especie_Tit & "'", DB, adOpenDynamic
'        NCod = RSCod!Codigo
'        RSCod.Close
        
'        DB.BeginTrans
'        DB.Execute ("Insert Into tblCaixa (Protocolo_Cartorio,CodAto,QRCode,N_Selo,Serie,Cod_Seg,Emolumentos,FRJ,FRC,NCod,TipoSelo) values ('" & RS!Protocolo_Cartorio & "','" & CodAto & "','" & Caminho & "','" & Codigo & "','" & Serie & "','" & CodSeguranca & "','" & Replace(Emol, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & NCod & "','" & Tipo & "')")
'        DB.Execute ("Update tblSeloDigital set Usado='" & 1 & "',Protocolo='" & RS!Protocolo_Cartorio & "',Data_Uso='" & Format(Date, "mm/dd/yyyy") & "',Ato='" & CodAto & "' Where IdSelo='" & RSSeloDigital!IdSelo & "'")
'        DB.CommitTrans
        RSSeloDigital.MoveNext
        Next
'<<< FIM PREPARA A TABELA tblCaixa >>>


'<<< ROTINA DE VERIFICAÇÃO DOS SELOS POSTECIPADOS >>>
    If RSPostecipa.State = 1 Then RSPostecipa.Close
    RSPostecipa.Open "Select * From tblSeloDigital Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Tipo='" & 113 & "' AND Postecipado='" & 0 & "'", DB, adOpenDynamic
    If RSPostecipa.RecordCount > 0 Then

        yato = RSPostecipa.RecordCount

        For xAto = 1 To yato
        '<<< Arredonda os centavos >>>
        Salva = 0
        Codigo = RSPostecipa!Codigo
        Serie = RSPostecipa!Serie
        Tipo = RSPostecipa!Tipo
        CodSeguranca = RSPostecipa!CodSeguranca
        'Protesto
'        If RSPostecipa!Ato >= 144 And RSPostecipa!Ato <= 151 Then
'            Emolumento = vProt
'            Emol = eProt
'            CodAto = RSPostecipa!Ato
'            FRC = frcProt
'            FRJ = frjProt
'            Codigo = RSPostecipa!Codigo
'            Serie = RSPostecipa!Serie
'            Tipo = RSPostecipa!Tipo
'            CodSeguranca = RSPostecipa!CodSeguranca
'        End If
        'Apontamento
        If RSPostecipa!Ato = 152 Or RSPostecipa!Ato = 756 Then
            Emolumento = vAponta
            Emol = eAponta
            CodAto = RSPostecipa!Ato
            FRC = frcAponta
            FRJ = frjAponta
            Codigo = RSPostecipa!Codigo
            Serie = RSPostecipa!Serie
            Tipo = RSPostecipa!Tipo
            CodSeguranca = RSPostecipa!CodSeguranca
            Salva = 1
            Ato3 = FormatCurrency(eAponta, 2)
        End If
        'Distribuidor
        If RSPostecipa!Ato = 179 Or RSPostecipa!Ato = 755 Then
            Emolumento = vDistrib
            Emol = eDistrib
            CodAto = RSPostecipa!Ato
            FRC = frcDistrib
            FRJ = frjDistrib
            Codigo = RSPostecipa!Codigo
            Serie = RSPostecipa!Serie
            Tipo = RSPostecipa!Tipo
            CodSeguranca = RSPostecipa!CodSeguranca
            Salva = 1
            Ato4 = FormatCurrency(eDistrib, 2)
        End If
        'Intimação
        If RSPostecipa!Ato = 162 Or RSPostecipa!Ato = 757 Then
            Emolumento = vIntima
            Emol = eIntima
            CodAto = RSPostecipa!Ato
            FRC = frcIntima
            FRJ = frjIntima
            Codigo = RSPostecipa!Codigo
            Serie = RSPostecipa!Serie
            Tipo = RSPostecipa!Tipo
            CodSeguranca = RSPostecipa!CodSeguranca
            Salva = 1
            Ato5 = FormatCurrency(eIntima, 2)
        End If
        'Edital
        If RSPostecipa!Ato = 164 Or RSPostecipa!Ato = 760 Then
            Emolumento = vEdital
            Emol = eEdital
            CodAto = RSPostecipa!Ato
            FRC = frcEdital
            FRJ = frjEdital
            Codigo = RSPostecipa!Codigo
            Serie = RSPostecipa!Serie
            Tipo = RSPostecipa!Tipo
            CodSeguranca = RSPostecipa!CodSeguranca
            Salva = 1
            Ato6 = FormatCurrency(eEdital, 2)
        End If
        'CPD
        If RSPostecipa!Ato = 180 Then
            Emolumento = vCpd
            Emol = eCpd
            CodAto = RSPostecipa!Ato
            FRC = frcCpd
            FRJ = frjCpd
            Codigo = RSPostecipa!Codigo
            Serie = RSPostecipa!Serie
            Tipo = RSPostecipa!Tipo
            CodSeguranca = RSPostecipa!CodSeguranca
            Salva = 1
            Ato1 = FormatCurrency(eCpd, 2)
        End If
        
        FRJ = Format(FRJ, "0.00")
        FRC = Format(FRC, "0.00")
        totFRJ = totFRJ + FRJ
        totFRC = totFRC + FRC


'        Value = "https://apps.tjpa.jus.br/ValidaSeloDigital/Detalhes/Resultado?codigoSelo=" & Codigo & "&serie=" & Serie & "&codigoTipoSelo=" & Tipo & "&codigoSeguranca=" & CodSeguranca & ""
'        Caminho = ("\\" & Server & "\QRCimg\Cx") & Tipo_Baixa & RS!Protocolo_Cartorio & "(" & xAto + 4 & ")" & ".bmp"
''        GenerateBMP StrPtr(Caminho), StrPtr(Value), 3, 5, QualityLow
'
'        RSCod.Open "Select * From tblEspecieTitulos WHERE Especie='" & RS!Especie_Tit & "'", DB, adOpenDynamic
'        NCod = RSCod!Codigo
'        RSCod.Close

'        If Salva = 1 Then
'            DB.BeginTrans
'            DB.Execute ("Insert Into tblCaixa (Protocolo_Cartorio,CodAto,QRCode,N_Selo,Serie,Cod_Seg,Emolumentos,FRJ,FRC,NCod,TipoSelo) values ('" & RS!Protocolo_Cartorio & "','" & CodAto & "','" & Caminho & "','" & Codigo & "','" & Serie & "','" & CodSeguranca & "','" & Replace(Emol, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & NCod & "','" & Tipo & "')")
'            DB.Execute ("Update tblSeloDigital set Postecipado='" & 1 & "',Data_Uso='" & Format(Date, "mm/dd/yyyy") & "' Where IdSelo='" & RSPostecipa!IdSelo & "'")
'            DB.CommitTrans
'        End If
        RSPostecipa.MoveNext
        Next

    End If
'<<< FIM DA ROTINA DE VERIFICAÇÃO DOS SELOS POSTECIPADOS >>>

         
'                DB.Execute ("Insert Into tblFinanceiro (CodPortador,Data_Ocorrencia,Protocolo,Custas,Tipo_Ocorrencia,Pagar,Valor_Tit,Valor_Juros,Valor_CPMF,Valor_Selo,Valor_Mora,V_Multa,Valor_Distrib,Data_Retirada,Devedor,Usuario,Hora,QtdeSG,QtdeSC,Pagante,Especie_Tit,Baixa_Lote) values ('" & RS!CodPortador & "','" & Format(Date, "ddmmyyyy") & "','" & RS!Protocolo_Cartorio & "','" & Replace(Custas, ",", ".") & "','" & 3 & "','" & Replace(Total, ",", ".") & "','" & Replace(Valor, ",", ".") & "','" & Replace(Juros, ",", ".") & "','" & Replace(CPMF, ",", ".") & "','" & Replace(Selo, ",", ".") & "','" & Replace(Mora, ",", ".") & "','" & Replace(Multa, ",", ".") & "','" & Replace(Distribuidor, ",", ".") & "','" & Format(Date, "mm/dd/yyyy") & "','" & RS!Devedor & "','" & User & "','" & Time & "','" & SeloGeral & "','" & SeloCertidao & "','" & cxNomeRecibo & "','" & RS!Especie_Tit & "','" & 1 & "')")
'                DB.Execute ("Update tblTitulo set Baixado='" & 1 & "',Valor_Custas='" & Replace(Custas, ",", ".") & "',Valor_Selo='" & Replace(Selo, ",", ".") & "',Valor_Distrib='" & Replace(Distribuidor, ",", ".") & "',Tipo_Ocorrencia='" & 3 & "',Data_Ocorrencia='" & Format(Date, "ddmmyyyy") & "',Data_Retirada='" & Format(Date, "mm/dd/yyyy") & "',CancelaBanco='" & CancelaBanco & "' Where Protocolo_Dist='" & RSBaixa!Protocolo & "'")
'
'            PrintConn
            
'            RSImp.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "'AND Data_Ocorrencia='" & Format(Date, "ddmmyyyy") & "'AND Estorno Is Null", DB, adOpenDynamic
'            If RSImp.RecordCount = 0 Then
'                MsgBox "Sem dados para impressão!", vbInformation
'                DB.RollbackTrans
'                Screen.MousePointer = 1
'                RSImp.Close
'                Exit Sub
'            End If
'                If RS!CodPortador > 0 Then
'                    RSBanco1.Open "Select * From tblBanco Where idBanco='" & RS!CodPortador & "'", DB, adOpenDynamic
'                    Portador = RSBanco1!Nome_Banco
'                    RSBanco1.Close
'                Else
'                    Portador = RS!Portador
'                End If
                
'Me.txtPagar = Custas
txtTotal = FormatCurrency(Total, 2)
'<<< F I M   R E T I R A D A >>>
    Ato9 = totFRJ
    Ato10 = totFRC
'    Valor = Custas + Selo + Distribuidor
'    Me.txtValorPagar = FormatCurrency(Total, 2)
'    Me.txtValorCartorio = FormatCurrency((Custas + Selo + Distribuidor), 2)
'    Me.txtTotal = Me.txtValorPagar
    
    Else
        MsgBox "O Título " & RSBaixa!Protocolo & " não pode ser RETIRADO.", vbInformation
    End If
        DB.Execute ("Insert Into tblGuiaProvisoria (N_Guia,Protocolo,DataEntrada,Ocorrencia,Pagar,Devedor,Custas,Selo,Valor_Titulo,Ato1,Ato2,Ato3,Ato4,Ato5,Ato6,Ato7,Ato8,Ato9,Ato10,Usuario,Num_Devedor,Tipo_Baixa) values ('" & 0 & "','" & txtProtocolo & "','" & Format(Date, "mm/dd/yyyy") & "','" & Me.txtOcorrencia & "','" & Replace(CDbl(txtTotal), ",", ".") & "','" & Trim(txtDevedor) & "','" & Replace(CDbl(Custas), ",", ".") & "','" & Replace(CDbl(Selo), ",", ".") & "','" & Replace(CDbl(Valor), ",", ".") & "','" & Replace(CDbl(Ato1), ",", ".") & "','" & Replace(CDbl(Ato2), ",", ".") & "','" & Replace(CDbl(Ato3), ",", ".") & "','" & Replace(CDbl(Ato4), ",", ".") & "','" & Replace(CDbl(Ato5), ",", ".") & "','" & Replace(CDbl(Ato6), ",", ".") & "','" & Replace(CDbl(Ato7), ",", ".") & "','" & Replace(CDbl(Ato8), ",", ".") & "','" & Replace(CDbl(totFRJ), ",", ".") & "','" & Replace(CDbl(totFRC), ",", ".") & "','" & User & "','" & RS!Num_Devedor & "','" & Tipo_Baixa & "')")
        DB.Execute ("Insert Into tblCalculo (Protocolo,Tipo_Ocorrencia,Valor,Devedor,Custas,Selo,Valor_Titulo,Ato1,Ato2,Ato3,Ato4,Ato5,Ato6,Ato7,Ato8,Ato9,Ato10,Usuario) values ('" & txtProtocolo & "','" & Tipo_Baixa & "','" & Replace(CDbl(txtTotal), ",", ".") & "','" & Trim(txtDevedor) & "','" & Replace(CDbl(Custas), ",", ".") & "','" & Replace(CDbl(Selo), ",", ".") & "','" & Replace(CDbl(Valor), ",", ".") & "','" & Replace(CDbl(Ato1), ",", ".") & "','" & Replace(CDbl(Ato2), ",", ".") & "','" & Replace(CDbl(Ato3), ",", ".") & "','" & Replace(CDbl(Ato4), ",", ".") & "','" & Replace(CDbl(Ato5), ",", ".") & "','" & Replace(CDbl(Ato6), ",", ".") & "','" & Replace(CDbl(Ato7), ",", ".") & "','" & Replace(CDbl(Ato8), ",", ".") & "','" & Replace(CDbl(Ato9), ",", ".") & "','" & Replace(CDbl(Ato10), ",", ".") & "','" & User & "')")
    End If
    
End If

Exit Sub

Erro:
    If Err.Number = -2147217873 Then
        Dim RSUser As ADODB.Recordset
        Set RSUser = New ADODB.Recordset

        RSUser.Open "SELECT * FROM tblGuiaProvisoria", DB, adOpenDynamic
        If RSUser.RecordCount > 0 Then
            MsgBox "O Protocolo n° " & RS!Protocolo_Cartorio & " já foi inserido pelo Usuário " & RSUser!Usuario, vbCritical
        End If
        RSUser.Close
    Else
        MsgBox "Erro de sistema n° " & Err.Number & "-" & Err.Description, vbCritical
    End If
End Sub
Public Function Calcula_Atos(RSAponta As ADODB.Recordset)
    eAponta = RSAponta!Apontamento
    frcAponta = Format(eAponta * 0.025, "0.00")
    frjAponta = Format(eAponta * 0.15, "0.00")
    vAponta = eAponta + frcAponta + frjAponta
    issAponta = Format(eAponta * 0.05, "0.00")
    
    eIntima = RSAponta!Intimacao
    frcIntima = Format(eIntima * 0.025, "0.00")
    frjIntima = Format(eIntima * 0.15, "0.00")
    vIntima = eIntima + frcIntima + frjIntima
    issIntima = Format(eIntima * 0.05, "0.00")
    
    eCanc = RSAponta!CancAponta
    frcCanc = Format(eCanc * 0.025, "0.00")
    frjCanc = Format(eCanc * 0.15, "0.00")
    vCanc = eCanc + frcCanc + frjCanc
    issCanc = Format(eCanc * 0.05, "0.00")
    
    eDistrib = RSAponta!Distribuidor
    frcDistrib = Format(eDistrib * 0.025, "0.00")
    frjDistrib = Format(eDistrib * 0.15, "0.00")
    vDistrib = eDistrib + frcDistrib + frjDistrib
    issDistrib = Format(eDistrib * 0.05, "0.00")
    
    eCpd = RSAponta!CPD
    frcCpd = Format(eCpd * 0.025, "0.00")
    frjCpd = Format(eCpd * 0.15, "0.00")
    vCpd = eCpd + frcCpd + frjCpd
    issCpd = Format(eCpd * 0.05, "0.00")
    
    eEdital = RSAponta!V_Edital
    frcEdital = Format(eEdital * 0.025, "0.00")
    frjEdital = Format(eEdital * 0.15, "0.00")
    vEdital = eEdital + frcEdital + frjEdital
    issEdital = Format(eEdital * 0.05, "0.00")
    
    eCtProt = RSAponta!ContraProtesto
    frcCtProt = Format(eCtProt * 0.025, "0.00")
    frjCtProt = Format(eCtProt * 0.15, "0.00")
    vCtProt = eCtProt + frcCtProt + frjCtProt
    issCtProt = Format(eCtProt * 0.05, "0.00")
End Function
Function SeparaExtenso(Extenso As String, iLineLenght As Integer) As String

Dim iCount As Integer
Dim X1 As String
Dim X2 As String

X1 = Left(Extenso, iLineLenght)
X2 = Mid(Extenso, iLineLenght + 1)

If Not Len(Extenso) <= iLineLenght Then
If Right(X1, 1) <> " " And Left(X2, 1) <> " " Then
For iCount = Len(Extenso) To 1 Step -1
If Mid(X1, iCount, 1) = " " Then
X1 = Left(Extenso, iCount)
X2 = Mid(Extenso, iCount + 1)
Exit For
End If
Next iCount
End If
End If
Texto1 = X1
Texto2 = X2

End Function
Public Sub Incluir()
On Error GoTo Erro

Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
Dim Cont As Integer

If Me.cboBaixa = "" Then
    MsgBox "Selecione o Tipo de Baixa.", vbInformation
    Exit Sub
End If


Limpar
MeuGridTitulos
Carrega_GridTitulos

Exit Sub
Erro:
    If Err.Number = -2147217873 Then
        RS.Open "SELECT * FROM tblGuiaProvisoria WHERE Protocolo ='" & Me.txtProtocolo & "'", DB, adOpenDynamic
        MsgBox "O título " & Me.txtProtocolo & " já foi inserido na Guia pelo Usuário " & RS!Usuario, vbInformation
    Else
        MsgBox "Erro de Sistema. " & Err.Description & " - N° " & Err.Number, vbCritical
    End If
    DB.RollbackTrans
    Limpar
End Sub

Public Sub Baixar()
On Error GoTo Erro

    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim RSBaixa As ADODB.Recordset
    Set RSBaixa = New ADODB.Recordset
    Dim RSAval As ADODB.Recordset
    Set RSAval = New ADODB.Recordset
    Dim RSBanco As ADODB.Recordset
    Set RSBanco = New ADODB.Recordset
    Dim RSAdianta As ADODB.Recordset
    Set RSAdianta = New ADODB.Recordset
    Dim nDataProt As Date
    Dim RSAponta As ADODB.Recordset
    Set RSAponta = New ADODB.Recordset
    Dim RSSeloDigital As ADODB.Recordset
    Set RSSeloDigital = New ADODB.Recordset
    Dim RSSeloGratuito As ADODB.Recordset
    Set RSSeloGratuito = New ADODB.Recordset
    Dim RSProtesto As ADODB.Recordset
    Set RSProtesto = New ADODB.Recordset
    Dim RSPostecipa As ADODB.Recordset
    Set RSPostecipa = New ADODB.Recordset
    Dim RSCod As ADODB.Recordset
    Set RSCod = New ADODB.Recordset
    Dim dOcorrencia As Date
    Dim RSImp As ADODB.Recordset
    Set RSImp = New ADODB.Recordset
    Dim Custas As Double
    Dim RSBanco1 As ADODB.Recordset
    Set RSBanco1 = New ADODB.Recordset
    Dim Portador As String
    Dim DataLiquida As Date
    
    FRJ = 0
    FRC = 0
    totFRJ = 0
    totFRC = 0
    Multa = 0
    CPMF = 0
    Juros = 0
    SeloGeral = 0
    SeloCertidao = 0
    Mora = 0
    Portador = ""
    CancelaBanco = 0
    Anuencia = False
    CodNota = 0
    N_Guia = ""
    Cartao = 0
    rProtocolo = 0
    txtValorCustas = 0
    CDeb = 0

'<<< Verifica Pasta >>>
    If Len(Dir("\\" & Server & "\ProtestoSCP\Caixa\" & Replace(Date, "/", ""), vbDirectory) & "") > 0 Then
    Else
        MkDir "\\" & Server & "\ProtestoSCP\Caixa\" & Replace(Date, "/", "")
    End If

    RSBaixa.Open "SELECT * FROM tblGuias WHERE N_Guia='" & Me.GridGuias.TextMatrix(x, 6) & "'AND Baixado='" & 0 & "'", DB, adOpenDynamic
    If RSBaixa.RecordCount = 0 Then
        MsgBox "Sem títulos selecionados para LIQUIDAR.", vbInformation
        Exit Sub
    End If

    If Me.optDataLiquida = True Then
        DataLiquida = InputBox("Digite a Data da Liquidação. (DD/MM/AAAA).")
    Else
        DataLiquida = Date
    End If

    Screen.MousePointer = 1
'    frmCDebito.Show 1
    nResp = MsgBox("Pagamento em DINHEIRO?", vbQuestion + vbYesNo)
    If nResp = vbNo Then
        CDeb = 1
    End If
        
    Tipo_Ocorrencia = RSBaixa!Tipo_Baixa
'<<< Abre Tabela de Apontamento >>>
    RSAponta.Open "Select * From tblApontamento", DB, adOpenDynamic
    
    Protocolo = ""
    
    rProtocolo = RSBaixa!Protocolo
    
    If Tipo_Ocorrencia = "CANCELAMENTO" Or Tipo_Ocorrencia = "Cancelamento" Then
    
    Do While Not RSBaixa.EOF
'<<< C A N C E L A M E N T O >>>

'<<< Busca o Protocolo na Tabela tblTitulo >>>
    RS.Open "Select * from tblTitulo where Protocolo_Cartorio=" & RSBaixa!Protocolo, DB, adOpenDynamic

    If RS.RecordCount = 0 Then
        MsgBox "O Título " & RSBaixa!Protocolo & " não pertence a este cartório.", vbInformation
        Exit Sub
    End If
    
        If RS!Tipo_Ocorrencia = "2" Then
            rProtocolo = RSBaixa!Protocolo
            Tipo_Baixa = "Cancelamento"
            Valor = RS!Saldo
            CalculoFaixas RSAponta
'            If Valor <= RSAponta!Faixa0 Then ValorCustas = FormatCurrency(RSAponta!Cancelado0, 2): vPago = RSAponta!sPago0: atoCanc = 154: nFaixa0 = 0
'            If Valor <= RSAponta!Faixa1 And Valor > RSAponta!Faixa0 Then ValorCustas = FormatCurrency(RSAponta!Cancelado1, 2): vPago = RSAponta!sPago1: atoCanc = 155
'            If Valor <= RSAponta!Faixa2 And Valor > RSAponta!Faixa1 Then ValorCustas = FormatCurrency(RSAponta!Cancelado2, 2): vPago = RSAponta!sPago2: atoCanc = 156
'            If Valor <= RSAponta!Faixa3 And Valor > RSAponta!Faixa2 Then ValorCustas = FormatCurrency(RSAponta!Cancelado3, 2): vPago = RSAponta!sPago3: atoCanc = 157
'            If Valor <= RSAponta!Faixa4 And Valor > RSAponta!Faixa3 Then ValorCustas = FormatCurrency(RSAponta!Cancelado4, 2): vPago = RSAponta!sPago4: atoCanc = 158
'            If Valor <= RSAponta!Faixa5 And Valor > RSAponta!Faixa4 Then ValorCustas = FormatCurrency(RSAponta!Cancelado5, 2): vPago = RSAponta!sPago5: atoCanc = 159
'            If Valor <= RSAponta!Faixa6 And Valor > RSAponta!Faixa5 Then ValorCustas = FormatCurrency(RSAponta!Cancelado6, 2): vPago = RSAponta!sPago6: atoCanc = 160
'            If Valor > RSAponta!Faixa6 Then ValorCustas = FormatCurrency(RSAponta!Cancelado7, 2): vPago = RSAponta!sPago7: atoCanc = 161
            ValorCustas = txtValorCustas
            Calcula_Atos RSAponta

            ePago = vPago
            frcPago = Format(ePago * 0.025, "0.00")
            frjPago = Format(ePago * 0.15, "0.00")
            vPago = ePago + frcPago + frjPago
            issPago = Format(ePago * 0.05, "0.00")
                                                          
            ISS = issPago + issCanc
            
            nTipoCancela = "A"
            Tipo_Baixa = "Cancelamento"
            Custas = ValorCustas - vCpd
            Selo = RSAponta!Selo * 2
'            Distribuidor = vDistrib
            Distribuidor = 0
            Total = Custas + Selo + ISS
            Juros = 0
            dOcorrencia = Date
            Adiantamento = FormatCurrency(0, 2)
                                    
            If RS!CodPortador > 0 Then
                RSBanco.Open "Select * From tblBanco Where idBanco='" & RS!CodPortador & "'", DB, adOpenDynamic
                Portador = RSBanco!Nome_Banco
                CodNota = RSBanco!CodNota
                Cobranca = RSBanco!Cobranca
                RSBanco.Close
            Else
                Portador = RS!Portador
            End If

'            If CodNota > 1 Then
'                RSSeloDigital.Open "Select * From tblSeloDigital Where Usado='" & 0 & "' AND Tipo='" & 102 & "' Order By IdSelo", DB, adOpenDynamic
'                Custas = 0
'                Total = 0
'                Selo = 0
'            Else
                RSSeloDigital.Open "Select * From tblSeloDigital Where Usado='" & 0 & "' AND Tipo='" & 101 & "' Order By IdSelo", DB, adOpenDynamic
'            End If
            
            If RSSeloDigital.RecordCount = 0 Then
                MsgBox "Sistema sem selos Digitais.", vbInformation, "Sem Selos"
                Screen.MousePointer = 1
                Exit Sub
            End If
            RSSeloDigital.MoveFirst
                    
            If RSAval.State = 1 Then RSAval.Close
            RSAval.Open "Select * From tblAvalista Where Protocolo_Cartorio='" & RSBaixa!Protocolo & "'", DB, adOpenDynamic
            nAvalista = RSAval.RecordCount
            
            Do While Not RSAval.EOF
                DB.Execute ("Update tblAvalista set Protestado='" & 0 & "' Where Protocolo_Cartorio='" & RSBaixa!Protocolo & "'")
                RSAval.MoveNext
            Loop
                                        
    '<<< PREPARA A TABELA tblCaixa >>>
        totFRC = 0
        totFRJ = 0
        totISS = 0
        
        RSPostecipa.Open "Select * From tblSeloDigital Where Protocolo='" & RSBaixa!Protocolo & "' AND Tipo='" & 113 & "'", DB, adOpenDynamic
        If RSPostecipa.RecordCount > 0 Then
            yato = 2
            RSPostecipa.Close
        End If
            For xAto = 1 To 2
            '<<< Arredonda os centavos >>>
            Select Case xAto
                Case 1
                    Emolumento = vPago
                    Emol = ePago
                    CodAto = atoCanc
                    FRC = frcPago
                    FRJ = frjPago
                Case 2
                    Emolumento = vCanc
                    Emol = eCanc
                    CodAto = 893
                    FRC = frcCanc
                    FRJ = frjCanc
'                Case 3
'                    Emolumento = vDistrib
'                    Emol = eDistrib
'                    CodAto = 179
'                    FRC = frcDistrib
'                    FRJ = frjDistrib
            End Select
            If RSSeloDigital!Tipo = "102" Then
                FRC = 0
                FRJ = 0
                Emol = 0
            End If
            
            FRJ = Format(FRJ, "0.00")
            FRC = Format(FRC, "0.00")
            totFRJ = totFRJ + FRJ
            totFRC = totFRC + FRC
            ISS = Format(Emol * 0.05, "0.00")
            totISS = totISS + ISS
               
            Codigo = RSSeloDigital!Codigo
            Serie = RSSeloDigital!Serie
            Tipo = RSSeloDigital!Tipo
            CodSeguranca = RSSeloDigital!CodSeguranca
                
            Value = "https://apps.tjpa.jus.br/ValidaSeloDigital/Detalhes/Resultado?codigoSelo=" & Codigo & "&serie=" & Serie & "&codigoTipoSelo=" & Tipo & "&codigoSeguranca=" & CodSeguranca & ""
            Caminho = ("\\" & Server & "\ProtestoSCP\QRCimg\Cx") & Tipo_Baixa & RS!Protocolo_Cartorio & "(" & xAto & ")" & ".bmp"
            GenerateBMP StrPtr(Caminho), StrPtr(Value), 3, 5, QualityLow
            
            RSCod.Open "Select * From tblEspecieTitulos WHERE Especie='" & RS!Especie_Tit & "'", DB, adOpenDynamic
            NCod = RSCod!Codigo
            RSCod.Close
            
            DB.BeginTrans
            DB.Execute ("Insert Into tblCaixa (Protocolo_Cartorio,CodAto,QRCode,N_Selo,Serie,Cod_Seg,Emolumentos,FRJ,FRC,NCod,TipoSelo,ISS) values ('" & RS!Protocolo_Cartorio & "','" & CodAto & "','" & Caminho & "','" & Codigo & "','" & Serie & "','" & CodSeguranca & "','" & Replace(Emol, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & NCod & "','" & Tipo & "','" & Replace(ISS, ",", ".") & "')")
            DB.Execute ("Update tblSeloDigital set Usado='" & 1 & "',Protocolo='" & RS!Protocolo_Cartorio & "',Data_Uso='" & Format(Date, "mm/dd/yyyy") & "',Ato='" & CodAto & "' Where IdSelo='" & RSSeloDigital!IdSelo & "'")
            DB.CommitTrans
            RSSeloDigital.MoveNext
            Next
    
    Taxa = 0
    vTaxa = 0
    
    If CDeb = 1 Then
        Cartao = 1
'        If opSafra = True Then
'            Taxa = (RSAponta!txSafra / 100) + 1
'        End If
'        If opBradesco = True Then
'            Taxa = (RSAponta!txBradesco / 100) + 1
'        End If
'        If opItau = True Then
'            Taxa = (RSAponta!txItau / 100) + 1
'        End If
'
'        vTaxa = FormatCurrency(Total * (Taxa - 1), 2)
'        Total = FormatCurrency(Total * Taxa, 2)
    End If

    If CDeb = 2 Then
        Codigo = 1
    Else
        Codigo = 0
    End If
    
'        nRespCENPROT = MsgBox("Cancelamento Presencial?", vbQuestion + vbYesNo, "Cancelamento")
        nRespCENPROT = vbYes
        If nRespCENPROT = vbYes Then
            Anuencia = True
        Else
            Anuencia = False
        End If
          
                    DB.Execute ("Insert Into tblFinanceiro (CodPortador,Data_Ocorrencia,Protocolo,Custas,Tipo_Ocorrencia,Pagar,Valor_Tit,Valor_Juros,Valor_CPMF,Valor_Selo,Valor_Mora,V_Multa,Valor_Distrib,Data_Cancelado,Devedor,Usuario,Hora,CancelaBanco,QtdeSG,QtdeSC,Pagante,Especie_Tit,TaxaCartao,Codigo,ISS,Cartao) values ('" & RS!CodPortador & "','" & Format(DataLiquida, "ddmmyyyy") & "','" & RS!Protocolo_Cartorio & "','" & Replace(Custas, ",", ".") & "','" & nTipoCancela & "','" & Replace(Total, ",", ".") & "','" & Replace(Valor, ",", ".") & "','" & Replace(Juros, ",", ".") & "','" & Replace(CPMF, ",", ".") & "','" & Replace(Selo, ",", ".") & "','" & Replace(Mora, ",", ".") & "','" & Replace(Multa, ",", ".") & "','" & 0 & "','" & Format(DataLiquida, "mm/dd/yyyy") & "','" & RS!Devedor & "','" & User & "','" & Time & "','" & CancelaBanco & "','" & SeloGeral & "','" & SeloCertidao & "','" & Portador & "','" & RS!Especie_Tit & _
                    "','" & Replace(vTaxa, ",", ".") & "','" & Codigo & "','" & Replace(totISS, ",", ".") & "','" & Cartao & "')")
                    DB.Execute ("Update tblTitulo set Baixado='" & 1 & "',Tipo_Ocorrencia='" & nTipoCancela & "',Data_Ocorrencia='" & Format(DataLiquida, "ddmmyyyy") & "',Data_Cancelado='" & Format(DataLiquida, "mm/dd/yyyy") & "',CancelaBanco='" & CancelaBanco & "',Anuencia='" & Anuencia & "' Where Protocolo_Cartorio='" & RSBaixa!Protocolo & "'")
                    
                PrintConn
                
                Data_Ocorrencia = Format(DataLiquida, "ddmmyyyy")

                RSImp.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Tipo_Ocorrencia='" & "A" & "'AND Data_Ocorrencia='" & Data_Ocorrencia & "'AND Estorno Is Null", DB, adOpenDynamic
                If RSImp.RecordCount = 0 Then
                    MsgBox "Sem dados para impressão!", vbInformation
                    DB.RollbackTrans
                    Screen.MousePointer = 1
                    Exit Sub
                End If
                    
                frmReciboCancelamento.Show 1
                    
                Set RSPrt = New ADODB.Recordset
                If RSPrt.State = 1 Then RSPrt.Close
                RSPrt.Open "Select * From tblCaixa", Conn
'                Set CrysRep = CrysApp.OpenReport(App.Path & "\Recibo_CaixaQRC.rpt")
                Set CrysRep = CrysApp.OpenReport(App.Path & "\Recibo_RetiradaQRC.rpt")
'                FileLoca = ("\\" & Server & "\ProtestoSCP\Caixa\" & Replace(Date, "/", "") & "\Cancelamento" & RS!Protocolo_Cartorio & "_" & Replace(Time, ":", "") & ".pdf")
                Dim caracteresInvalidos As String
                Dim formatDevedor As String
                Dim i As Integer
                
                ' Definindo os caracteres inválidos
                caracteresInvalidos = "'*´""'"
                
                formatDevedor = RS!Devedor
                
                ' Substituindo por espaço
                For i = 1 To Len(caracteresInvalidos)
                    formatDevedor = Replace(formatDevedor, Mid(caracteresInvalidos, i, 1), " ")
                Next i
                
                FileLoca = ("\\" & Server & "\ProtestoSCP\Caixa\" & Replace(Date, "/", "") & "\" & UCase(Replace(Trim(formatDevedor), "/", "")) & RS!Protocolo_Cartorio & "_CCL_" & Right(Replace(Time, ":", ""), 2) & ".pdf")
                
                If Not RSPrt.EOF Then
                    With CrysRep
                        Call .Database.Tables(1).SetDataSource(RSPrt)
                        If RS!CodPortador > 0 Then
                            RSBanco.Open "Select * From tblBanco Where idBanco='" & RS!CodPortador & "'", DB, adOpenDynamic
                            Portador = RSBanco!Nome_Banco
                            RSBanco.Close
                        Else
                            Portador = RS!Portador
                        End If
                        .DiscardSavedData
                        .EnableParameterPrompting = False
                        .ReadRecords
                        .ParameterFields(1).AddCurrentValue "Recibo de " & Tipo_Baixa & " " & FormatCurrency(RSImp!Pagar, 2)
                        .ParameterFields(2).AddCurrentValue "Protocolo: " & RSImp!Protocolo
'                        .ParameterFields(3).AddCurrentValue "Recebemos de: " & UCase(Portador)
                        .ParameterFields(3).AddCurrentValue "Recebemos de: " & UCase(cxNomeRecibo)
                        Texto = SeparaExtenso(Extenso(RSImp!Pagar), 40)
                        .ParameterFields(4).AddCurrentValue "A importância de: " & Texto1 & Texto2
                        .ParameterFields(5).AddCurrentValue "Referentes as custas de: " & Tipo_Baixa & ", do Apontamento e do Registro do Protesto"
                        .ParameterFields(6).AddCurrentValue "Do Titulo num. " & Trim(RS!Num_Titulo)
                        .ParameterFields(7).AddCurrentValue "Vencido em " & Format(RS!Vencimento, "00/00/0000")
                        .ParameterFields(8).AddCurrentValue "Sacador " & RS!Sacador
                        .ParameterFields(9).AddCurrentValue "Cedente " & RS!Cedente
                        .ParameterFields(10).AddCurrentValue "Apresentado por " & UCase(Portador)
                        .ParameterFields(11).AddCurrentValue "Contra :" & RS!Devedor & " Conforme discriminacao abaixo."
                        .ParameterFields(12).AddCurrentValue "Entrada " & RS!Data_Apresenta
                        .ParameterFields(13).AddCurrentValue "Nosso Numero " & RS!Nosso_Num
                        .ParameterFields(14).AddCurrentValue "Valor do Titulo: " & FormatCurrency(RSImp!Valor_Tit, 2)
                        If RSImp!Valor_Juros > 0 Then
                            .ParameterFields(15).AddCurrentValue "Juros: " & FormatCurrency(RSImp!Valor_Juros, 2)
                        End If
                        If CDeb = 1 Then
                            .ParameterFields(15).AddCurrentValue "Desp. Cartão de Débito: " & FormatCurrency(RSImp!Pagar - RSImp!Custas - RSImp!Valor_Selo - RSImp!ISS, 2)
                        End If
                        If CDeb = 2 Then
                            .ParameterFields(15).AddCurrentValue "Desp. Boleto Bancário: " & FormatCurrency(RSImp!TaxaCartao, 2)
                        End If
                        
                                            
'                        nData = CDate("01/12/2019")
'                        If RS!Data_Apresenta < nData Then
                            .ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(totFRJ, 2)
'                        Else
'                            .ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(totFRJ - frjDistrib, 2)
'                            totFRJ = totFRJ - frjDistrib
'                        End If
'                        If RS!Data_Apresenta < nData Then
                            .ParameterFields(17).AddCurrentValue "Taxa FRC: " & FormatCurrency(totFRC, 2)
'                        Else
'                            .ParameterFields(17).AddCurrentValue "Taxa FRC: " & FormatCurrency(totFRC - frcDistrib, 2)
'                            totFRC = totFRC - frcDistrib
'                        End If
                        DB.BeginTrans
                        DB.Execute ("Update tblFinanceiro set FRJ='" & Replace(totFRJ, ",", ".") & "',FRC='" & Replace(totFRC, ",", ".") & "' Where idFinanceiro='" & RSImp!idFinanceiro & "'")
                        DB.CommitTrans
                        If RSImp!Valor_Selo > 0 Then
                            .ParameterFields(18).AddCurrentValue "Selo Judicial: " & FormatCurrency(RSImp!Valor_Selo, 2)
                        End If
'                        If RSImp!Valor_Distrib > 0 Then
'                            .ParameterFields(19).AddCurrentValue "Distribuidor: " & FormatCurrency(eDistrib, 2)
'                        End If
                        .ParameterFields(20).AddCurrentValue "Desp. Canc. Apont.: " & FormatCurrency(eCanc, 2)
                        'TT = FormatCurrency(RSImp!Custas - (Aponta!CancAponta), 2)
                        .ParameterFields(21).AddCurrentValue "Desp. Cancelamento: " & FormatCurrency(ePago, 2)
'                        If CDeb = 1 Then
'                            .ParameterFields(22).AddCurrentValue "Desp. Cartão de Débito: " & FormatCurrency(RSImp!Pagar - RSImp!Custas - RSImp!Valor_Selo - RSImp!ISS, 2)
'                        End If
'                        If CDeb = 2 Then
'                            .ParameterFields(22).AddCurrentValue "Desp. Boleto Bancário: " & FormatCurrency(RSImp!TaxaCartao, 2)
'                        End If
                        .ParameterFields(23).AddCurrentValue ""
                        'If Me.chkEditado = True Then
                            .ParameterFields(24).AddCurrentValue ""
                        '    TT = FormatCurrency(RSImp!Custas - (Aponta!Apontamento + Aponta!Intimacao + Aponta!CancAponta + Aponta!CPD + Aponta!V_Edital), 2)
                        'Else
                        '    TT = FormatCurrency(RSImp!Custas - (Aponta!CancAponta), 2)
                        'End If
                        .ParameterFields(25).AddCurrentValue ""
                        .ParameterFields(26).AddCurrentValue "TOTAL PAGO       :" & FormatCurrency(RSImp!Pagar, 2)
                        .ParameterFields(27).AddCurrentValue "" & "Belém, " & Day(CDate(Format(Data_Ocorrencia, "00/00/0000"))) & " de " & MonthName(Month(CDate(Format(Data_Ocorrencia, "00/00/0000")))) & " de " & Year(CDate(Format(Data_Ocorrencia, "00/00/0000")))
                        .ParameterFields(28).AddCurrentValue "Protocolo Dist.: " & RS!Protocolo_Dist
                        .ParameterFields(29).AddCurrentValue "ISS: " & RSImp!ISS
    '                    .ParameterFields(28).AddCurrentValue "Autenticacao: " & RSImp!Protocolo * (Day(Date) & Month(Date) & Year(Date))
                        '.ParameterFields(29).AddCurrentValue "" & NumSelo1 & "-" & NumSelo2 & "-" & NumSelo3 & "-" & NumSelo4 & "-" & NumSelo5 & "-" & NumSelo6 & "-" & NumSelo7
                    End With
                    Set CRExportOptions = CrysRep.ExportOptions
                    CRExportOptions.FormatType = crEFTPortableDocFormat
                    CRExportOptions.DestinationType = crEDTDiskFile
                    CRExportOptions.DiskFileName = FileLoca
                    CrysRep.DisplayProgressDialog = False
                    CrysRep.Export False
                    Set CRExportOptions = Nothing
                End If
                nResp = MsgBox("Imprimir o Recibo de Cancelamento?", vbQuestion + vbYesNo)
                If nResp = vbYes Then
                    CrysRep.PrintOut False, 1
                End If
                    
                Screen.MousePointer = 1
'                frmPrtCaixa.CRViewer1.ReportSource = CrysRep
'                frmPrtCaixa.CRViewer1.ViewReport
'                frmPrtCaixa.Show 1
                If Conn.State = 1 Then Conn.Close

                CaixaXML
'                EnvioXML

            '<<< Apaga QRCode >>>
                RSPrt.Open "Select * From tblCaixa", DB, adOpenDynamic
                    Do While Not RSPrt.EOF
                        Kill RSPrt!QRCode
                        RSPrt.MoveNext
                    Loop
                RSPrt.Close

                DB.Execute "Delete From tblCaixa"
                DB.Execute ("UPDATE tblGuias SET Baixado='" & 1 & "'WHERE N_Guia='" & RSBaixa!N_Guia & "'")
'<<< F I M   C A N C E L A M E N T O >>>

'<<< C U S T A S  D E  P R O T E S T O >>>
If RS!Data_Apresenta < CDate("01/12/2019") And RS!Especie_Tit <> "CDA" Then
Else
'    If CodNota > 1 Then
'    Else
    rProtocolo = RSBaixa!Protocolo
    Tipo_Baixa = "Protesto"
    Valor = RS!Saldo
    CalculoFaixaProtesto RSAponta
'        If Valor <= RSAponta!Faixa0 Then ValorCustas = FormatCurrency(RSAponta!Pago720, 2): vPago = RSAponta!sPago0: atoPago = 171: vProt = RSAponta!sProt0: atoCanc = 154: atoProt = 144: nFaixa0 = 0
'        If Valor <= RSAponta!Faixa1 And Valor > RSAponta!Faixa0 Then ValorCustas = FormatCurrency(RSAponta!Pago721, 2): vPago = RSAponta!sPago1: atoPago = 172: vProt = RSAponta!sProt1: atoCanc = 155: atoProt = 145
'        If Valor <= RSAponta!Faixa2 And Valor > RSAponta!Faixa1 Then ValorCustas = FormatCurrency(RSAponta!Pago722, 2): vPago = RSAponta!sPago2: atoPago = 173: vProt = RSAponta!sProt2: atoCanc = 156: atoProt = 146
'        If Valor <= RSAponta!Faixa3 And Valor > RSAponta!Faixa2 Then ValorCustas = FormatCurrency(RSAponta!Pago723, 2): vPago = RSAponta!sPago3: atoPago = 174: vProt = RSAponta!sProt3: atoCanc = 157: atoProt = 147
'        If Valor <= RSAponta!Faixa4 And Valor > RSAponta!Faixa3 Then ValorCustas = FormatCurrency(RSAponta!Pago724, 2): vPago = RSAponta!sPago4: atoPago = 175: vProt = RSAponta!sProt4: atoCanc = 158: atoProt = 148
'        If Valor <= RSAponta!Faixa5 And Valor > RSAponta!Faixa4 Then ValorCustas = FormatCurrency(RSAponta!Pago725, 2): vPago = RSAponta!sPago5: atoPago = 176: vProt = RSAponta!sProt5: atoCanc = 159: atoProt = 149
'        If Valor <= RSAponta!Faixa6 And Valor > RSAponta!Faixa5 Then ValorCustas = FormatCurrency(RSAponta!Pago726, 2): vPago = RSAponta!sPago6: atoPago = 177: vProt = RSAponta!sProt6: atoCanc = 160: atoProt = 150
'        If Valor > RSAponta!Faixa6 Then ValorCustas = FormatCurrency(RSAponta!Pago727, 2): vPago = RSAponta!sPago7: atoPago = 178: vProt = RSAponta!sProt7: atoCanc = 161: atoProt = 151

        ePago = vPago
        frcPago = Format(ePago * 0.025, "0.00")
        frjPago = Format(ePago * 0.15, "0.00")
        vPago = ePago + frcPago + frjPago
        issPago = Format(ePago * 0.05, "0.00")
        
        eProt = vProt
        frcProt = Format(eProt * 0.025, "0.00")
        frjProt = Format(eProt * 0.15, "0.00")
        vProt = eProt + frcProt + frjProt
        issProt = Format(eProt * 0.05, "0.00")
        
        If RS!Editado = False Or IsNull(RS!Editado) Then
            issEdital = 0
        End If
        
        If RS!ContraProtesto = False Or IsNull(RS!ContraProtesto) Then
            issCtProt = 0
        End If
        
        ISS = issProt + issAponta + issCpd + issIntima + issEdital + issCtProt + issDistrib
        
        If RS!Editado = True Then
            nAtos = 5
            Custas = FormatCurrency(vProt + vAponta + vCpd + vIntima + vEdital, 2)
        Else
            nAtos = 4
            Custas = FormatCurrency(vProt + vAponta + vCpd + vIntima, 2)
        End If
        
        If RS!ContraProtesto = True And RS!Editado = True Then
            nAtos = 6
            ContraProtesto = 1
            Custas = FormatCurrency(vProt + vAponta + vCpd + vIntima + vEdital + vCtProt, 2)
        End If

        If RS!ContraProtesto = True And RS!Editado = False Then
            nAtos = 5
            ContraProtesto = 1
            Custas = FormatCurrency(vProt + vAponta + vCpd + vIntima + vCtProt, 2)
        End If
    
    'Verifica se tem Avalista
    RSAval.Close
    RSAval.Open "Select * From tblAvalista Where Protocolo_Cartorio='" & RS!Protocolo_Cartorio & "'", DB, adOpenDynamic
    
    Do While Not RSAval.EOF
        If RSAval!Editado = True Then
            nAvalEdital = nAvalEdital + 1
            Custas = Custas + vEdital
            nAtos = nAtos + 1
        End If
        If RSAval!Intimado = True Then
            nAvalIntima = nAvalIntima + 1
            Custas = Custas + vIntima
            nAtos = nAtos + 1
        End If
        nAvalista = nAvalista + 1
        RSAval.MoveNext
    Loop
    RSAval.Close
    'FIM Verifica se tem Avalista
    
    Selo = (nAtos + 1) * RSAponta!Selo
    Distribuidor = vDistrib
    Total = Custas + Selo + Distribuidor + ISS
    Tipo_Baixa = "Protesto"
    
    RSAdianta.Open "Select * From tblAdiantamento Where Protocolo_Cartorio='" & RS!Protocolo_Cartorio & "'", DB, adOpenDynamic
    If IsNull(RS!Adiantamento) Then
        Adiantamento = 0
    Else
        Adiantamento = RS!Adiantamento
    End If
    If RSAdianta.RecordCount = 0 Then
        'Adiantamento = FormatCurrency(RS!Adiantamento, 2)
        DB.Execute ("Insert Into tblAdiantamento (Protocolo_Cartorio,Devedor,Valor,Saldo,Sacador,Cedente,Adiantamento,Baixa,DataPagamento,Data_Entrada) values ('" & RS!Protocolo_Cartorio & "','" & RS!Devedor & "','" & Replace(RS!Valor, ",", ".") & "','" & Replace(RS!Saldo, ",", ".") & "','" & RS!Sacador & "','" & RS!Cedente & "','" & Replace(Adiantamento, ",", ".") & "','" & 1 & "','" & Format(Date, "mm/dd/yyyy") & "','" & Format(RS!Data_Apresenta, "mm/dd/yyyy") & "')")
    Else
        'Adiantamento = FormatCurrency(RSAdianta!Adiantamento, 2)
        DB.Execute ("Update tblAdiantamento set Baixa='" & 1 & "',DataPagamento='" & Format(Date, "mm/dd/yyyy") & "' Where Protocolo_Cartorio='" & RS!Protocolo_Cartorio & "'")
    End If
    
    If RS!Fora_Area = True Then
        Custas = Custas
    Else
        Custas = Custas - Adiantamento
    End If
                    
    Codigo = 0
    
    If CDeb = 1 Then
        Cartao = 1
        If opSafra = True Then
            Taxa = (RSAponta!txSafra / 100) + 1
        End If
        If opBradesco = True Then
            Taxa = (RSAponta!txBradesco / 100) + 1
        End If
        If opItau = True Then
            Taxa = (RSAponta!txItau / 100) + 1
        End If
'        Me.Label17.Visible = True
'        Me.Label17.Caption = "C. Débito"
'        frmCaixatxtValorCPMF.Visible = True
'        txtValorCartorio = FormatCurrency((Custas + Selo) * (Taxa), 2)
'        txtValorPagar = FormatCurrency(Total * (Taxa), 2)
'        vTaxa = Format(Total * (Taxa - 1), "0.00")
'        Total = Format(Total * (Taxa), "0.00")
        
    End If
    
    
'    If CodNota > 0 Then
'        DB.Execute ("Update tblTitulo set CustasProtesto='" & 1 & "' Where Protocolo_Cartorio='" & RS!Protocolo_Cartorio & "'")
''        DB.Execute ("Update tblTitulo set CustasProtesto='" & 1 & "' Where Protocolo_Dist='" & RS!Protocolo_Cartorio & "'")
'        DB.Execute ("Insert Into tblFinanceiro (CodPortador,Data_Ocorrencia,Protocolo,Custas,Tipo_Ocorrencia,Pagar,Valor_Tit,Valor_Juros,Valor_CPMF,Valor_Selo,Valor_Mora,V_Multa,Valor_Distrib,Devedor,Usuario,Hora,QtdeSG,QtdeSC,Pgto_CH,Pagante,Especie_Tit,TaxaCartao,Codigo,ISS,Cartao) values ('" & RS!CodPortador & "','" & Format(DataLiquida, "ddmmyyyy") & "','" & RS!Protocolo_Cartorio & "','" & Replace(Custas, ",", ".") & "','" & 1 & "','" & Replace(Total, ",", ".") & "','" & Replace(Valor, ",", ".") & "','" & Replace(Juros, ",", ".") & "','" & Replace(CPMF, ",", ".") & "','" & Replace(Selo, ",", ".") & "','" & Replace(Mora, ",", ".") & "','" & Replace(Multa, ",", ".") & "','" & Replace(Distribuidor, ",", ".") & "','" & RS!Devedor & "','" & User & "','" & Time & "','" & SeloGeral & "','" & SeloCertidao & "','" & vCheque & "','" & Portador & "','" & RS!Especie_Tit & "','" & Replace(vTaxa, ",", ".") & "','" & Codigo & "','" & Replace(ISS, ",", ".") & "','" & Cartao & "')")
'    Else
        DB.Execute ("Update tblTitulo set Data_Pagamento='" & Format(DataLiquida, "mm/dd/yyyy") & "',CustasProtesto='" & 1 & "' Where Protocolo_Cartorio='" & RS!Protocolo_Cartorio & "'")
'        DB.Execute ("Update tblTitulo set Data_Pagamento='" & Format(Date, "mm/dd/yyyy") & "',CustasProtesto='" & 1 & "' Where Protocolo_Dist='" & RS!Protocolo_Cartorio & "'")
        DB.Execute ("Insert Into tblFinanceiro (CodPortador,Data_Ocorrencia,Protocolo,Custas,Tipo_Ocorrencia,Pagar,Valor_Tit,Valor_Juros,Valor_CPMF,Valor_Selo,Valor_Mora,V_Multa,Valor_Distrib,Data_Pagamento,Devedor,Usuario,Hora,QtdeSG,QtdeSC,Pgto_CH,Pagante,Especie_Tit,TaxaCartao,Codigo,ISS,Cartao) values ('" & RS!CodPortador & "','" & Format(DataLiquida, "ddmmyyyy") & "','" & RS!Protocolo_Cartorio & "','" & Replace(Custas, ",", ".") & "','" & 1 & "','" & Replace(Total, ",", ".") & "','" & Replace(Valor, ",", ".") & "','" & Replace(Juros, ",", ".") & "','" & Replace(CPMF, ",", ".") & "','" & Replace(Selo, ",", ".") & "','" & Replace(Mora, ",", ".") & "','" & Replace(Multa, ",", ".") & "','" & Replace(Distribuidor, ",", ".") & "','" & Format(DataLiquida, "mm/dd/yyyy") & "','" & RS!Devedor & "','" & User & "','" & Time & "','" & SeloGeral & "','" & SeloCertidao & "','" & 0 & "','" & Portador & "','" & RS!Especie_Tit & "','" & _
        Replace(vTaxa, ",", ".") & "','" & Codigo & "','" & Replace(ISS, ",", ".") & "','" & Cartao & "')")
'    End If

'<<< PREPARA A TABELA tblCaixa >>>
'    If Not IsNull(RS!CodPortador) Then
'        If RS!CodPortador = 0 Then
'            Portador = RS!Portador
'        Else
'            RSBanco.Open "Select * from tblBanco Where idBanco = '" & RS!CodPortador & "'", DB, adOpenDynamic
'            Portador = RSBanco!Nome_Banco
'            RSBanco.Close
'        End If
'    End If
        
    yato = 1
    If RS!ContraProtesto = True Then
        yato = 2
        ContraProtesto = 1
    End If
    
    totFRJ = 0
    totFRC = 0
    
    If RSPostecipa.State = 1 Then RSPostecipa.Close
    RSPostecipa.Open "Select * From tblSeloDigital Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Tipo='" & 113 & "' AND Ato='" & 180 & "'", DB, adOpenDynamic
    If RSPostecipa.RecordCount = 0 Then
    For xAto = 1 To yato
    '<<< Arredonda os centavos >>>
            
    Select Case xAto
        Case 1
            Emolumento = vCpd
            Emol = eCpd
            CodAto = 967
            FRC = frcCpd
            FRJ = frjCpd

        Case 2
            Emolumento = vCtProt
            Emol = eCtProt
            CodAto = 966
            FRC = frcCtProt
            FRJ = frjCtProt
    End Select


    FRJ = Format(FRJ, "0.00")
    FRC = Format(FRC, "0.00")
    totFRJ = totFRJ + FRJ
    totFRC = totFRC + FRC
    ISS = Format(Emol * 0.05, "0.00")
    totISS = totISS + ISS
    
    Codigo = RSSeloDigital!Codigo
    Serie = RSSeloDigital!Serie
    Tipo = RSSeloDigital!Tipo
    CodSeguranca = RSSeloDigital!CodSeguranca
        
    Value = "https://apps.tjpa.jus.br/ValidaSeloDigital/Detalhes/Resultado?codigoSelo=" & Codigo & "&serie=" & Serie & "&codigoTipoSelo=" & Tipo & "&codigoSeguranca=" & CodSeguranca & ""
    Caminho = ("\\" & Server & "\ProtestoSCP\QRCimg\Cx") & Tipo_Baixa & RS!Protocolo_Cartorio & "(" & xAto & ")" & ".bmp"
    GenerateBMP StrPtr(Caminho), StrPtr(Value), 3, 5, QualityLow
    
    RSCod.Open "Select * From tblEspecieTitulos WHERE Especie='" & RS!Especie_Tit & "'", DB, adOpenDynamic
    NCod = RSCod!Codigo
    RSCod.Close
    
    DB.BeginTrans
    DB.Execute ("Insert Into tblCaixa (Protocolo_Cartorio,CodAto,QRCode,N_Selo,Serie,Cod_Seg,Emolumentos,FRJ,FRC,NCod,TipoSelo,ISS) values ('" & RS!Protocolo_Cartorio & "','" & CodAto & "','" & Caminho & "','" & Codigo & "','" & Serie & "','" & CodSeguranca & "','" & Replace(Emol, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & NCod & "','" & Tipo & "','" & Replace(ISS, ",", ".") & "')")
    DB.Execute ("Update tblSeloDigital set Usado='" & 1 & "',Protocolo='" & RS!Protocolo_Cartorio & "',Data_Uso='" & Format(Date, "mm/dd/yyyy") & "',Ato='" & CodAto & "' Where IdSelo='" & RSSeloDigital!IdSelo & "'")
    DB.CommitTrans
    RSSeloDigital.MoveNext
    Next
End If
        
'<<< ROTINA DE VERIFICAÇÃO DOS SELOS POSTECIPADOS >>>
    If RSPostecipa.State = 1 Then RSPostecipa.Close
    RSPostecipa.Open "Select * From tblSeloDigital Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Tipo='" & 113 & "' AND Postecipado='" & 0 & "'", DB, adOpenDynamic
    If RSPostecipa.RecordCount > 0 Then
    
        yato = RSPostecipa.RecordCount
    
        For xAto = 1 To yato
        '<<< Arredonda os centavos >>>
        Salva = 0
        Codigo = RSPostecipa!Codigo
        Serie = RSPostecipa!Serie
        Tipo = RSPostecipa!Tipo
        CodSeguranca = RSPostecipa!CodSeguranca
        'Protesto
        If RSPostecipa!Data_Uso1 >= "2024-03-11" And RSPostecipa!Data_Uso1 <= "2024-03-15" Then
            If RSPostecipa!Ato >= 894 And RSPostecipa!Ato <= 959 Then
                Emolumento = vProt
                Emol = eProt
                CodAto = atoCanc
                FRC = frcProt
                FRJ = frjProt
                Codigo = RSPostecipa!Codigo
                Serie = RSPostecipa!Serie
                Tipo = RSPostecipa!Tipo
                CodSeguranca = RSPostecipa!CodSeguranca
                Salva = 1
            End If
        Else
            If RSPostecipa!Ato >= 144 And RSPostecipa!Ato <= 151 Or RSPostecipa!Ato >= 827 And RSPostecipa!Ato <= 892 Then
                Emolumento = vProt
                Emol = eProt
                CodAto = atoProt
                FRC = frcProt
                FRJ = frjProt
                Codigo = RSPostecipa!Codigo
                Serie = RSPostecipa!Serie
                Tipo = RSPostecipa!Tipo
                CodSeguranca = RSPostecipa!CodSeguranca
                Salva = 1
            End If
        End If
'        If RSPostecipa!Ato >= 144 And RSPostecipa!Ato <= 151 Or RSPostecipa!Ato >= 827 And RSPostecipa!Ato <= 892 Then
'            Emolumento = vProt
'            Emol = eProt
'            CodAto = atoProt
'            FRC = frcProt
'            FRJ = frjProt
'            Codigo = RSPostecipa!Codigo
'            Serie = RSPostecipa!Serie
'            Tipo = RSPostecipa!Tipo
'            CodSeguranca = RSPostecipa!CodSeguranca
'            Salva = 1
'        End If
        'Apontamento
        If RSPostecipa!Ato = 152 Or RSPostecipa!Ato = 756 Then
            Emolumento = vAponta
            Emol = eAponta
            CodAto = 756
            FRC = frcAponta
            FRJ = frjAponta
            Codigo = RSPostecipa!Codigo
            Serie = RSPostecipa!Serie
            Tipo = RSPostecipa!Tipo
            CodSeguranca = RSPostecipa!CodSeguranca
            Salva = 1
        End If
        'Distribuidor
        If RSPostecipa!Ato = 179 Or RSPostecipa!Ato = 755 Then
            Emolumento = vDistrib
            Emol = eDistrib
            CodAto = 755
            FRC = frcDistrib
            FRJ = frjDistrib
            Codigo = RSPostecipa!Codigo
            Serie = RSPostecipa!Serie
            Tipo = RSPostecipa!Tipo
            CodSeguranca = RSPostecipa!CodSeguranca
            Salva = 1
        End If
        'Intimação
        If RSPostecipa!Ato = 162 Or RSPostecipa!Ato = 757 Then
            Emolumento = vIntima
            Emol = eIntima
            CodAto = 757
            FRC = frcIntima
            FRJ = frjIntima
            Codigo = RSPostecipa!Codigo
            Serie = RSPostecipa!Serie
            Tipo = RSPostecipa!Tipo
            CodSeguranca = RSPostecipa!CodSeguranca
            Salva = 1
        End If
        'Edital
        If RSPostecipa!Ato = 164 Or RSPostecipa!Ato = 760 Then
            Emolumento = vEdital
            Emol = eEdital
            CodAto = 760
            FRC = frcEdital
            FRJ = frjEdital
            Codigo = RSPostecipa!Codigo
            Serie = RSPostecipa!Serie
            Tipo = RSPostecipa!Tipo
            CodSeguranca = RSPostecipa!CodSeguranca
            Salva = 1
        End If
        'CPD
        If RSPostecipa!Ato = 180 Then
            Emolumento = vCpd
            Emol = eCpd
            CodAto = 967
            FRC = frcCpd
            FRJ = frjCpd
            Codigo = RSPostecipa!Codigo
            Serie = RSPostecipa!Serie
            Tipo = RSPostecipa!Tipo
            CodSeguranca = RSPostecipa!CodSeguranca
            Salva = 1
        End If
        
        FRJ = Format(FRJ, "0.00")
        FRC = Format(FRC, "0.00")
        totFRJ = totFRJ + FRJ
        totFRC = totFRC + FRC
        ISS = Format(Emol * 0.05, "0.00")
        totISS = totISS + ISS

        Value = "https://apps.tjpa.jus.br/ValidaSeloDigital/Detalhes/Resultado?codigoSelo=" & Codigo & "&serie=" & Serie & "&codigoTipoSelo=" & Tipo & "&codigoSeguranca=" & CodSeguranca & ""
        Caminho = ("\\" & Server & "\ProtestoSCP\QRCimg\Cx") & Tipo_Baixa & RS!Protocolo_Cartorio & "(" & xAto + 2 & ")" & ".bmp"
        GenerateBMP StrPtr(Caminho), StrPtr(Value), 3, 5, QualityLow
        
        RSCod.Open "Select * From tblEspecieTitulos WHERE Especie='" & RS!Especie_Tit & "'", DB, adOpenDynamic
        NCod = RSCod!Codigo
        RSCod.Close
        
        If Salva = 1 Then
            DB.BeginTrans
            DB.Execute ("Insert Into tblCaixa (Protocolo_Cartorio,CodAto,QRCode,N_Selo,Serie,Cod_Seg,Emolumentos,FRJ,FRC,NCod,TipoSelo,ISS) values ('" & RS!Protocolo_Cartorio & "','" & CodAto & "','" & Caminho & "','" & Codigo & "','" & Serie & "','" & CodSeguranca & "','" & Replace(Emol, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & NCod & "','" & Tipo & "','" & Replace(ISS, ",", ".") & "')")
            DB.Execute ("Update tblSeloDigital set Postecipado='" & 1 & "',Data_Uso='" & Format(Date, "mm/dd/yyyy") & "' Where IdSelo='" & RSPostecipa!IdSelo & "'")
            DB.CommitTrans
            RSPostecipa.MoveNext
        End If
        Next

    End If
'<<< FIM PREPARA A TABELA tblCaixa >>>
        
     PrintConn
     If RSImp.State = 1 Then RSImp.Close
     RSImp.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Tipo_Ocorrencia='" & 1 & "'AND Data_Ocorrencia='" & Data_Ocorrencia & "'AND Estorno Is Null", DB, adOpenDynamic
     If RSImp.RecordCount = 0 Then
         MsgBox "Sem dados para impressão!", vbInformation
         DB.RollbackTrans
         Screen.MousePointer = 1
         RSImp.Close
         Exit Sub
     End If
                
     Set RSPrt = New ADODB.Recordset
     If RSPrt.State = 1 Then RSPrt.Close
     RSProtesto.Open "Select * From tblProtestoCopia Where Protocolo_Cartorio='" & RS!Protocolo_Cartorio & "'", DB, adOpenDynamic
     RSPrt.Open "Select * From tblCaixa", Conn
     Set CrysRep = CrysApp.OpenReport(App.Path & "\Recibo_RetiradaQRC.rpt")
'     FileLoca = ("\\" & Server & "\Caixa\" & Replace(Date, "/", "") & "\CustaProtesto" & RS!Protocolo_Cartorio & "_" & Replace(Time, ":", "") & ".pdf")

     FileLoca = ("\\" & Server & "\ProtestoSCP\Caixa\" & Replace(Date, "/", "") & "\" & UCase(Trim(Replace(formatDevedor, "/", ""))) & RS!Protocolo_Cartorio & "_CPT_" & Right(Replace(Time, ":", ""), 2) & ".pdf")
     If Not RSPrt.EOF Then
         With CrysRep
             Call .Database.Tables(1).SetDataSource(RSPrt)
             .DiscardSavedData
             .EnableParameterPrompting = False
             .ReadRecords
             If RSAdianta.RecordCount = 0 Then
                 vTotal = FormatCurrency(RSImp!Pagar, 2) - FormatCurrency(Adiantamento, 2)
             Else
                 vTotal = FormatCurrency(RSImp!Pagar, 2) - FormatCurrency(RSAdianta!Adiantamento, 2)
             End If
             .ParameterFields(1).AddCurrentValue "Recibo de " & Tipo_Baixa & " " & FormatCurrency(RSImp!Pagar, 2)
             .ParameterFields(2).AddCurrentValue "Protocolo: " & RSImp!Protocolo
             .ParameterFields(3).AddCurrentValue "Recebemos de: " & UCase(cxNomeRecibo)
             Texto = SeparaExtenso(Extenso(RSImp!Pagar), 40)
             .ParameterFields(4).AddCurrentValue "A importância de: " & Texto1 & Texto2
             .ParameterFields(5).AddCurrentValue "Referentes as custas de Certidão de Protesto "
             .ParameterFields(6).AddCurrentValue "Do Titulo num. " & RS!Num_Titulo
             .ParameterFields(7).AddCurrentValue "Vencido em " & Format(RS!Vencimento, "00/00/0000")
             .ParameterFields(8).AddCurrentValue "Sacador " & RS!Sacador
             .ParameterFields(9).AddCurrentValue "Cedente " & RS!Cedente
             If RS!CodPortador > 0 Then
                 RSBanco.Open "Select * From tblBanco Where idBanco='" & RS!CodPortador & "'", DB, adOpenDynamic
                 .ParameterFields(10).AddCurrentValue "Apresentado por " & RSBanco!Nome_Banco
                 RSBanco.Close
             Else
                 .ParameterFields(10).AddCurrentValue "Apresentado por " & RS!Portador
             End If
             .ParameterFields(11).AddCurrentValue "Contra " & RS!Devedor & " Conforme discriminação abaixo."
             .ParameterFields(12).AddCurrentValue "Entrada " & RS!Data_Apresenta
             .ParameterFields(13).AddCurrentValue "Nosso Numero " & RS!Nosso_Num
             .ParameterFields(14).AddCurrentValue "Valor do Titulo   :" & FormatCurrency(RSImp!Valor_Tit, 2)
             If RSImp!Valor_Juros > 0 Then
                 .ParameterFields(15).AddCurrentValue "Juros             :" & FormatCurrency(RSImp!Valor_Juros, 2)
             End If
             
             If CDeb = 1 Then
'                 .ParameterFields(15).AddCurrentValue "Desp. Cartão de Débito: " & FormatCurrency(RSImp!Pagar - RSImp!Custas - Distribuidor - Selo - RSImp!ISS, 2)
                 .ParameterFields(15).AddCurrentValue "Desp. Boleto Bancário: " & FormatCurrency(RSImp!TaxaCartao, 2)
             End If
            If CDeb = 2 Then
                .ParameterFields(15).AddCurrentValue "Desp. Boleto Bancário: " & FormatCurrency(RSImp!TaxaCartao, 2)
            End If
             
             If RS!Editado = False Or IsNull(RS!Editado) Then
                 '.ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(RSProtesto!txFRJ, 2)
                 frjSoma = frjProt + frjAponta + frjCpd + frjIntima + (frjIntima * nAvalista) + frjDistrib + (frjEdital * nAvalEdital)
                 frcSoma = frcProt + frcAponta + frcCpd + frcIntima + (frcIntima * nAvalista) + frcDistrib + (frcEdital * nAvalEdital)
                 .ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(frjSoma, 2)
                 .ParameterFields(17).AddCurrentValue "Taxa FRC: " & FormatCurrency(frcSoma, 2)
             Else
                 frjSoma = frjProt + frjAponta + frjCpd + frjIntima + (frjIntima * nAvalista) + frjDistrib + frjEdital + (frjEdital * nAvalEdital)
                 frcSoma = frcProt + frcAponta + frcCpd + frcIntima + (frcIntima * nAvalista) + frcDistrib + frcEdital + (frcEdital * nAvalEdital)
                 .ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(frjSoma, 2)
                 .ParameterFields(17).AddCurrentValue "Taxa FRC: " & FormatCurrency(frcSoma, 2)
             End If
             DB.BeginTrans
             DB.Execute ("Update tblFinanceiro set FRJ='" & Replace(frjSoma, ",", ".") & "',FRC='" & Replace(frcSoma, ",", ".") & "' Where idFinanceiro='" & RSImp!idFinanceiro & "'")
             DB.CommitTrans
             SomaProt = eProt
             If RS!ContraProtesto = True Then
                 SomaProt = eProt + eCtProt
                 frjSoma = frjSoma + frjCtProt
                 frcSoma = frcSoma + frcCtProt
                 .ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(frjSoma, 2)
                 .ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(frcSoma, 2)
             End If
             .ParameterFields(18).AddCurrentValue "Selo Judicial: " & FormatCurrency(Selo, 2)

'                If IsNull(RSProtesto!txFRJ) Then
'                    .ParameterFields(19).AddCurrentValue "Distribuidor: " & FormatCurrency(vDistrib, 2)
'                Else
                 .ParameterFields(19).AddCurrentValue "Distribuidor: " & FormatCurrency(eDistrib, 2)
'                End If
             .ParameterFields(20).AddCurrentValue "Desp. Apontamento: " & FormatCurrency(eAponta, 2)
             .ParameterFields(21).AddCurrentValue "Desp. Intimação: " & FormatCurrency(eIntima + (eIntima * nAvalista), 2)
             
             .ParameterFields(22).AddCurrentValue "Desp. CPD: " & FormatCurrency(eCpd, 2)
                 
             If RS!Editado = True And nAvalEdital > 0 Then
                 .ParameterFields(23).AddCurrentValue "Edital: " & FormatCurrency(eEdital + (eEdital * nAvalEdital), 2)
             End If
             
             If RS!Editado = True And nAvalEdital = 0 Then
                 .ParameterFields(23).AddCurrentValue "Edital: " & FormatCurrency(eEdital, 2)
             End If
             
             'TT = FormatCurrency(RS!V_Protesto, 2)
             .ParameterFields(24).AddCurrentValue "Desp. Protesto: " & FormatCurrency(SomaProt, 2)
             .ParameterFields(25).AddCurrentValue "Adiantamento -: " & FormatCurrency(Adiantamento, 2)
             .ParameterFields(26).AddCurrentValue "TOTAL PAGO: " & FormatCurrency(vTotal, 2)
             .ParameterFields(27).AddCurrentValue "" & "Belém, " & Day(CDate(Format(Data_Ocorrencia, "00/00/0000"))) & " de " & MonthName(Month(CDate(Format(Data_Ocorrencia, "00/00/0000")))) & " de " & Year(CDate(Format(Data_Ocorrencia, "00/00/0000")))
             .ParameterFields(28).AddCurrentValue "Protocolo Dist.: " & RS!Protocolo_Dist
             .ParameterFields(29).AddCurrentValue "ISS: " & RSImp!ISS
             '.ParameterFields(28).AddCurrentValue "Autenticacao: " & RSImp!Protocolo * (Day(Date) & Month(Date) & Year(Date))
             '.ParameterFields(29).AddCurrentValue "" & NumSelo1 & "-" & NumSelo2 & "-" & NumSelo3 & "-" & NumSelo4 & "-" & NumSelo5 & "-" & NumSelo6 & "-" & NumSelo7
         End With
         Set CRExportOptions = CrysRep.ExportOptions
         CRExportOptions.FormatType = crEFTPortableDocFormat
         CRExportOptions.DestinationType = crEDTDiskFile
         CRExportOptions.DiskFileName = FileLoca
         CrysRep.DisplayProgressDialog = False
         CrysRep.Export False
         Set CRExportOptions = Nothing

     End If

     RSProtesto.Close
    nResp = MsgBox("Imprimir o Recibo de Custas de Protesto?", vbQuestion + vbYesNo)
    If nResp = vbYes Then
        CrysRep.PrintOut False, 1
    End If
    Screen.MousePointer = 1

'     frmPrtCaixa.CRViewer1.ReportSource = CrysRep
'     frmPrtCaixa.CRViewer1.ViewReport
'     frmPrtCaixa.Show 1

     If Conn.State = 1 Then Conn.Close
        
     CaixaXML
     
'     EnvioXML
'<<< Apaga QRCode >>>
    RSPrt.Open "Select * From tblCaixa", DB, adOpenDynamic
        Do While Not RSPrt.EOF
            Kill RSPrt!QRCode
            RSPrt.MoveNext
        Loop
    RSPrt.Close
     
     DB.Execute "Delete From tblCaixa"
              
'    End If
End If
'<<< F I M  C U S T A S  D E  P R O T E S T O >>>

    DB.Execute ("UPDATE tblGuias SET Baixado='" & 1 & "'WHERE N_Guia='" & RSBaixa!N_Guia & "'")
    Else
        MsgBox "O Título " & RSBaixa!Protocolo & " não pode ser CANCELADO.", vbInformation
    End If
    If RS.State = 1 Then RS.Close
    If RSSeloDigital.State = 1 Then RSSeloDigital.Close
    If RSPostecipa.State = 1 Then RSPostecipa.Close
    If RSImp.State = 1 Then RSImp.Close
    If RSAdianta.State = 1 Then RSAdianta.Close
    RSBaixa.MoveNext
    Loop
'    MeuGrid
'    Carrega_Grid
'    MsgBox "Liquidação finalizada com sucesso!", vbInformation
        
    End If
    
    If Tipo_Ocorrencia = "RETIRADA" Or Tipo_Ocorrencia = "Retirada" Then
    
        FRJ = 0
        FRC = 0
        totFRJ = 0
        totFRC = 0
        Multa = 0
        CPMF = 0
        Juros = 0
        SeloGeral = 0
        SeloCertidao = 0
        Mora = 0
        cxNomeRecibo = ""
        CancelaBanco = 0
    
        Do While Not RSBaixa.EOF
        RS.Open "Select * from tblTitulo where Protocolo_Cartorio=" & RSBaixa!Protocolo, DB, adOpenDynamic
    
        If RS.RecordCount = 0 Then
            MsgBox "O Título " & RSBaixa!Protocolo & " não pertence a este cartório.", vbInformation
            Libera = 1
            Exit Sub
        End If
    
        If RS!Tipo_Ocorrencia = "" Or RS!Tipo_Ocorrencia = "0" Or IsNull(RS!Tipo_Ocorrencia) Then
            rProtocolo = RSBaixa!Protocolo
            Tipo_Baixa = "Retirada"
            Valor = RS!Saldo
            CalculoFaixas RSAponta
'            If Valor <= RSAponta!Faixa0 Then ValorCustas = FormatCurrency(RSAponta!Retirado0, 2): vRet = RSAponta!sPago0: atoPago = 171: nFaixa0 = 0
'            If Valor <= RSAponta!Faixa1 And Valor > RSAponta!Faixa0 Then ValorCustas = FormatCurrency(RSAponta!Retirado1, 2): vRet = RSAponta!sPago1: atoPago = 172
'            If Valor <= RSAponta!Faixa2 And Valor > RSAponta!Faixa1 Then ValorCustas = FormatCurrency(RSAponta!Retirado2, 2): vRet = RSAponta!sPago2: atoPago = 173
'            If Valor <= RSAponta!Faixa3 And Valor > RSAponta!Faixa2 Then ValorCustas = FormatCurrency(RSAponta!Retirado3, 2): vRet = RSAponta!sPago3: atoPago = 174
'            If Valor <= RSAponta!Faixa4 And Valor > RSAponta!Faixa3 Then ValorCustas = FormatCurrency(RSAponta!Retirado4, 2): vRet = RSAponta!sPago4: atoPago = 175
'            If Valor <= RSAponta!Faixa5 And Valor > RSAponta!Faixa4 Then ValorCustas = FormatCurrency(RSAponta!Retirado5, 2): vRet = RSAponta!sPago5: atoPago = 176
'            If Valor <= RSAponta!Faixa6 And Valor > RSAponta!Faixa5 Then ValorCustas = FormatCurrency(RSAponta!Retirado6, 2): vRet = RSAponta!sPago6: atoPago = 177
'            If Valor > RSAponta!Faixa6 Then ValorCustas = FormatCurrency(RSAponta!Retirado7, 2): vRet = RSAponta!sPago7: atoPago = 178
                                         
            ValorCustas = txtValorCustas
            eRet = vRet
            frcRet = Format(eRet * 0.025, "0.00")
            frjRet = Format(eRet * 0.15, "0.00")
            vRet = eRet + frcRet + frjRet
            issRet = Format(eRet * 0.05, "0.00")
            
            Calcula_Atos RSAponta
            
            If RS!Editado = False Or IsNull(RS!Editado) Then
                issEdital = 0
            End If
            
            ISS = issRet + issCpd + issAponta + issDistrib + issIntima + issEdital + issCanc
            
            Custas = ValorCustas
            Selo = RSAponta!Selo * 6
            Distribuidor = vDistrib
            Total = Custas + Selo + Distribuidor + ISS
            Juros = 0
            dOcorrencia = Date
            Adiantamento = FormatCurrency(0, 2)
            RSSeloDigital.Open "Select * From tblSeloDigital Where Usado='" & 0 & "' AND Tipo='" & 101 & "' Order By IdSelo", DB, adOpenDynamic
            
            If RSSeloDigital.RecordCount = 0 Then
                MsgBox "Sistema sem selos Digitais.", vbInformation, "Sem Selos"
                Screen.MousePointer = 1
                Exit Sub
            End If
            RSSeloDigital.MoveFirst
                
                
'<<< PREPARA A TABELA tblCaixa >>>
            totFRC = 0
            totFRJ = 0
            totISS = 0
'    <<< Busca Selo Digital >>>
            If RS!Especie_Tit = "CDA" Then
                RSBanco1.Open "Select * From tblBanco Where idBanco='" & RS!CodPortador & "'", DB, adOpenDynamic
                Portador = RSBanco1!Nome_Banco
                Cobranca = RSBanco1!Cobranca
                RSBanco1.Close
                If Cobranca = False Then
                    yato = 1
                    RSSeloDigital.Close
                    RSSeloDigital.Open "Select * From tblSeloDigital Where Usado='" & 0 & "' AND Tipo='" & 102 & "' Order By IdSelo", DB, adOpenDynamic
                    If RSSeloDigital.RecordCount = 0 Then
                        MsgBox "Sistema sem selos Digitais.", vbInformation, "Sem Selos"
                        Screen.MousePointer = 1
                        Exit Sub
                    End If
                    RSSeloDigital.MoveFirst
                Else
                    yato = 3
                    If RS!Editado = True Then
                        yato = yato + 1
                    End If
                    RSPostecipa.Open "Select * From tblSeloDigital Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Tipo='" & 113 & "' AND Postecipado='" & 0 & "'", DB, adOpenDynamic
                    If RSPostecipa.RecordCount > 0 Then
                        yato = 3
                        RSPostecipa.Close
                    End If
                End If
            Else
                yato = 3
                If RS!Editado = True Then
                    yato = yato + 1
                End If
                RSPostecipa.Open "Select * From tblSeloDigital Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Tipo='" & 113 & "' AND Postecipado='" & 0 & "'", DB, adOpenDynamic
                If RSPostecipa.RecordCount > 0 Then
                    Do While Not RSPostecipa.EOF
                        If RSPostecipa!Ato = 180 Then
                            yato = 2
                        End If
                        RSPostecipa.MoveNext
                    Loop
'                    yato = 3
                    RSPostecipa.Close
                End If
            End If
            For xAto = 1 To yato
            '<<< Arredonda os centavos >>>
            Select Case xAto
                Case 1
                    If RS!Especie_Tit = "CDA" And Cobranca = False Then
                        Emolumento = 0
                        Emol = 0
                        CodAto = atoPago
                        FRC = 0
                        FRJ = 0
                        Total = 0
                    Else
                        Emolumento = vRet
                        Emol = eRet
                        CodAto = atoPago
                        FRC = frcRet
                        FRJ = frjRet
                    End If
                Case 2
                    Emolumento = vCanc
                    Emol = eCanc
                    CodAto = 893
                    FRC = frcCanc
                    FRJ = frjCanc
                Case 3
                    Emolumento = vCpd
                    Emol = eCpd
                    CodAto = 967
                    FRC = frcCpd
                    FRJ = frjCpd
                Case 4
                    Emolumento = vEdital
                    Emol = eEdital
                    CodAto = 760
                    FRC = frcEdital
                    FRJ = frjEdital
            End Select
    
            FRJ = Format(FRJ, "0.00")
            FRC = Format(FRC, "0.00")
            totFRJ = totFRJ + FRJ
            totFRC = totFRC + FRC
            ISS = Format(Emol * 0.05, "0.00")
            totISS = totISS + ISS
    
            Codigo = RSSeloDigital!Codigo
            Serie = RSSeloDigital!Serie
            Tipo = RSSeloDigital!Tipo
            CodSeguranca = RSSeloDigital!CodSeguranca
                
            Value = "https://apps.tjpa.jus.br/ValidaSeloDigital/Detalhes/Resultado?codigoSelo=" & Codigo & "&serie=" & Serie & "&codigoTipoSelo=" & Tipo & "&codigoSeguranca=" & CodSeguranca & ""
            Caminho = ("\\" & Server & "\ProtestoSCP\QRCimg\Cx") & Tipo_Baixa & RS!Protocolo_Cartorio & "(" & xAto & ")" & ".bmp"
            GenerateBMP StrPtr(Caminho), StrPtr(Value), 3, 5, QualityLow
            
            RSCod.Open "Select * From tblEspecieTitulos WHERE Especie='" & RS!Especie_Tit & "'", DB, adOpenDynamic
            NCod = RSCod!Codigo
            RSCod.Close
            
            DB.BeginTrans
            DB.Execute ("Insert Into tblCaixa (Protocolo_Cartorio,CodAto,QRCode,N_Selo,Serie,Cod_Seg,Emolumentos,FRJ,FRC,NCod,TipoSelo,ISS) values ('" & RS!Protocolo_Cartorio & "','" & CodAto & "','" & Caminho & "','" & Codigo & "','" & Serie & "','" & CodSeguranca & "','" & Replace(Emol, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & NCod & "','" & Tipo & "','" & Replace(ISS, ",", ".") & "')")
            DB.Execute ("Update tblSeloDigital set Usado='" & 1 & "',Protocolo='" & RS!Protocolo_Cartorio & "',Data_Uso='" & Format(Date, "mm/dd/yyyy") & "',Ato='" & CodAto & "' Where IdSelo='" & RSSeloDigital!IdSelo & "'")
            DB.CommitTrans
            RSSeloDigital.MoveNext
            Next
'<<< FIM PREPARA A TABELA tblCaixa >>>


'<<< ROTINA DE VERIFICAÇÃO DOS SELOS POSTECIPADOS >>>
            If RSPostecipa.State = 1 Then RSPostecipa.Close
            RSPostecipa.Open "Select * From tblSeloDigital Where Protocolo='" & RS!Protocolo_Cartorio & "' AND Tipo='" & 113 & "' AND Postecipado='" & 0 & "'", DB, adOpenDynamic
            If RSPostecipa.RecordCount > 0 Then

                yato = RSPostecipa.RecordCount
        
                For xAto = 1 To yato
                '<<< Arredonda os centavos >>>
                Salva = 0
                Codigo = RSPostecipa!Codigo
                Serie = RSPostecipa!Serie
                Tipo = RSPostecipa!Tipo
                CodSeguranca = RSPostecipa!CodSeguranca
        
                'Apontamento
                If RSPostecipa!Ato = 152 Or RSPostecipa!Ato = 756 Then
                    Emolumento = vAponta
                    Emol = eAponta
                    CodAto = 756
                    FRC = frcAponta
                    FRJ = frjAponta
                    Codigo = RSPostecipa!Codigo
                    Serie = RSPostecipa!Serie
                    Tipo = RSPostecipa!Tipo
                    CodSeguranca = RSPostecipa!CodSeguranca
                    Salva = 1
                End If
                'Distribuidor
                If RSPostecipa!Ato = 179 Or RSPostecipa!Ato = 755 Then
                    Emolumento = vDistrib
                    Emol = eDistrib
                    CodAto = 755
                    FRC = frcDistrib
                    FRJ = frjDistrib
                    Codigo = RSPostecipa!Codigo
                    Serie = RSPostecipa!Serie
                    Tipo = RSPostecipa!Tipo
                    CodSeguranca = RSPostecipa!CodSeguranca
                    Salva = 1
                End If
                'Intimação
                If RSPostecipa!Ato = 162 Or RSPostecipa!Ato = 757 Then
                    Emolumento = vIntima
                    Emol = eIntima
                    CodAto = 757
                    FRC = frcIntima
                    FRJ = frjIntima
                    Codigo = RSPostecipa!Codigo
                    Serie = RSPostecipa!Serie
                    Tipo = RSPostecipa!Tipo
                    CodSeguranca = RSPostecipa!CodSeguranca
                    Salva = 1
                End If
                'Edital
                If RSPostecipa!Ato = 164 Or RSPostecipa!Ato = 760 Then
                    Emolumento = vEdital
                    Emol = eEdital
                    CodAto = 760
                    FRC = frcEdital
                    FRJ = frjEdital
                    Codigo = RSPostecipa!Codigo
                    Serie = RSPostecipa!Serie
                    Tipo = RSPostecipa!Tipo
                    CodSeguranca = RSPostecipa!CodSeguranca
                    Salva = 1
                End If
                'CPD
                If RSPostecipa!Ato = 180 Then
                    Emolumento = vCpd
                    Emol = eCpd
                    CodAto = 967
                    FRC = frcCpd
                    FRJ = frjCpd
                    Codigo = RSPostecipa!Codigo
                    Serie = RSPostecipa!Serie
                    Tipo = RSPostecipa!Tipo
                    CodSeguranca = RSPostecipa!CodSeguranca
                    Salva = 1
                End If
                
                FRJ = Format(FRJ, "0.00")
                FRC = Format(FRC, "0.00")
                totFRJ = totFRJ + FRJ
                totFRC = totFRC + FRC
                ISS = Format(Emol * 0.05, "0.00")
                totISS = totISS + ISS
        
        
                Value = "https://apps.tjpa.jus.br/ValidaSeloDigital/Detalhes/Resultado?codigoSelo=" & Codigo & "&serie=" & Serie & "&codigoTipoSelo=" & Tipo & "&codigoSeguranca=" & CodSeguranca & ""
                Caminho = ("\\" & Server & "\ProtestoSCP\QRCimg\Cx") & Tipo_Baixa & RS!Protocolo_Cartorio & "(" & xAto + 4 & ")" & ".bmp"
                GenerateBMP StrPtr(Caminho), StrPtr(Value), 3, 5, QualityLow
                
                RSCod.Open "Select * From tblEspecieTitulos WHERE Especie='" & RS!Especie_Tit & "'", DB, adOpenDynamic
                NCod = RSCod!Codigo
                RSCod.Close

                If Salva = 1 Then
                    DB.BeginTrans
                    DB.Execute ("Insert Into tblCaixa (Protocolo_Cartorio,CodAto,QRCode,N_Selo,Serie,Cod_Seg,Emolumentos,FRJ,FRC,NCod,TipoSelo,ISS) values ('" & RS!Protocolo_Cartorio & "','" & CodAto & "','" & Caminho & "','" & Codigo & "','" & Serie & "','" & CodSeguranca & "','" & Replace(Emol, ",", ".") & "','" & Replace(FRJ, ",", ".") & "','" & Replace(FRC, ",", ".") & "','" & NCod & "','" & Tipo & "','" & Replace(ISS, ",", ".") & "')")
                    DB.Execute ("Update tblSeloDigital set Postecipado='" & 1 & "',Data_Uso='" & Format(Date, "mm/dd/yyyy") & "' Where IdSelo='" & RSPostecipa!IdSelo & "'")
                    DB.CommitTrans
                End If
                RSPostecipa.MoveNext
                Next
            End If
'<<< FIM DA ROTINA DE VERIFICAÇÃO DOS SELOS POSTECIPADOS >>>
    Taxa = 0
    vTaxa = 0
    
    If CDeb = 1 Then
        Cartao = 1
'        If opSafra = True Then
'            Taxa = (RSAponta!txSafra / 100) + 1
'        End If
'        If opBradesco = True Then
'            Taxa = (RSAponta!txBradesco / 100) + 1
'        End If
'        If opItau = True Then
'            Taxa = (RSAponta!txItau / 100) + 1
'        End If
'
'        vTaxa = FormatCurrency(Total * (Taxa - 1), 2)
'        Total = FormatCurrency(Total * Taxa, 2)
    End If

    If CDeb = 2 Then
        Codigo = 1
    Else
        Codigo = 0
    End If
            DB.Execute ("Insert Into tblFinanceiro (CodPortador,Data_Ocorrencia,Protocolo,Custas,Tipo_Ocorrencia,Pagar,Valor_Tit,Valor_Juros,Valor_CPMF,Valor_Selo,Valor_Mora,V_Multa,Valor_Distrib,Data_Retirada,Devedor,Usuario,Hora,QtdeSG,QtdeSC,Pagante,Especie_Tit,TaxaCartao,ISS,Cartao) values ('" & RS!CodPortador & "','" & Format(DataLiquida, "ddmmyyyy") & "','" & RS!Protocolo_Cartorio & "','" & Replace(Custas, ",", ".") & "','" & 3 & "','" & Replace(Total, ",", ".") & "','" & Replace(Valor, ",", ".") & "','" & Replace(Juros, ",", ".") & "','" & Replace(CPMF, ",", ".") & "','" & Replace(Selo, ",", ".") & "','" & Replace(Mora, ",", ".") & "','" & Replace(Multa, ",", ".") & "','" & Replace(Distribuidor, ",", ".") & "','" & Format(DataLiquida, "mm/dd/yyyy") & "','" & RS!Devedor & "','" & User & "','" & Time & "','" & SeloGeral & "','" & SeloCertidao & "','" & cxNomeRecibo & "','" & RS!Especie_Tit & "','" & Replace(vTaxa, ",", ".") & "','" & Replace(totISS, ",", ".") & "','" & Cartao & "')")
            DB.Execute ("Update tblTitulo set Baixado='" & 1 & "',Valor_Custas='" & Replace(Custas, ",", ".") & "',Valor_Selo='" & Replace(Selo, ",", ".") & "',Valor_Distrib='" & Replace(Distribuidor, ",", ".") & "',Tipo_Ocorrencia='" & 3 & "',Data_Ocorrencia='" & Format(DataLiquida, "ddmmyyyy") & "',Data_Retirada='" & Format(DataLiquida, "mm/dd/yyyy") & "',CancelaBanco='" & CancelaBanco & "' Where Protocolo_Cartorio='" & RSBaixa!Protocolo & "'")
            
            PrintConn
            
            RSImp.Open "Select * From tblFinanceiro Where Protocolo='" & RS!Protocolo_Cartorio & "'AND Data_Ocorrencia='" & Format(DataLiquida, "ddmmyyyy") & "'AND Estorno Is Null", DB, adOpenDynamic
            If RSImp.RecordCount = 0 Then
                MsgBox "Sem dados para impressão!", vbInformation
                DB.RollbackTrans
                Screen.MousePointer = 1
                RSImp.Close
                Exit Sub
            End If
            If RS!CodPortador > 0 Then
                RSBanco1.Open "Select * From tblBanco Where idBanco='" & RS!CodPortador & "'", DB, adOpenDynamic
                Portador = RSBanco1!Nome_Banco
                RSBanco1.Close
            Else
                Portador = RS!Portador
            End If
        
            frmReciboCancelamento.Show 1
        
            Set RSPrt = New ADODB.Recordset
            If RSPrt.State = 1 Then RSPrt.Close
            RSPrt.Open "Select * From tblCaixa", Conn
            
            Set CrysRep = CrysApp.OpenReport(App.Path & "\Recibo_RetiradaQRC.rpt")
            FileLoca = ("\\" & Server & "\ProtestoSCP\Caixa\" & Replace(Date, "/", "") & "\" & UCase(Trim(Replace(RS!Devedor, "/", ""))) & RS!Protocolo_Cartorio & "_RET_" & Right(Replace(Time, ":", ""), 2) & ".pdf")
            If Not RSPrt.EOF Then
                With CrysRep
                    Call .Database.Tables(1).SetDataSource(RSPrt)
                    .DiscardSavedData
                    .EnableParameterPrompting = False
                    .ReadRecords
                    .ParameterFields(1).AddCurrentValue "Recibo de RETIRADA " & FormatCurrency(RSImp!Pagar, 2)
                    .ParameterFields(2).AddCurrentValue "Protocolo: " & RSImp!Protocolo
                    .ParameterFields(3).AddCurrentValue "Recebemos de: " & UCase(cxNomeRecibo)
                    Texto = SeparaExtenso(Extenso(RSImp!Pagar), 40)
                    .ParameterFields(4).AddCurrentValue "A importância de: " & Texto1 & Texto2 & " REFERENTES AS CUSTA DE RETIRADA SEM PROTESTO DO TÍTULO Nº " & RS!Num_Titulo & " NO VALOR DE " & FormatCurrency(RS!Saldo, 2) & " VENCIDO EM " & Mid(RS!Vencimento, 1, 2) & "/" & Mid(RS!Vencimento, 3, 2) & "/" & Right(RS!Vencimento, 4) & " NOSSO NÚMERO " & RS!Nosso_Num
                    .ParameterFields(9).AddCurrentValue "Cedente " & RS!Cedente
                    .ParameterFields(10).AddCurrentValue "Contra " & RSImp!Devedor & " Conforme discriminacao abaixo."
                    If RSImp!Valor_Juros > 0 Then
                        .ParameterFields(15).AddCurrentValue "Juros: " & FormatCurrency(RSImp!Valor_Juros, 2)
                    End If
                    If CDeb = 1 Then
'                        .ParameterFields(15).AddCurrentValue "Desp. Cartão de Débito: " & FormatCurrency(RSImp!Pagar - RSImp!Custas - RSImp!Valor_Selo - RSImp!ISS, 2)
                        .ParameterFields(15).AddCurrentValue "Desp. Boleto Bancário: " & FormatCurrency(RSImp!TaxaCartao, 2)
                    End If
                    If CDeb = 2 Then
                        .ParameterFields(15).AddCurrentValue "Desp. Boleto Bancário: " & FormatCurrency(RSImp!TaxaCartao, 2)
                    End If
                    
                    nData = CDate("01/12/2019")
                    If RS!Data_Apresenta > nData Then
                        .ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(totFRJ, 2)
                    Else
                        .ParameterFields(16).AddCurrentValue "Taxa FRJ: " & FormatCurrency(totFRJ + frjAponta + frjIntima + frjDistrib, 2)
                        totFRJ = totFRJ + frjAponta + frjIntima + frjDistrib
                    End If
                    If RS!Data_Apresenta > nData Then
                        .ParameterFields(17).AddCurrentValue "Taxa FRC: " & FormatCurrency(totFRC, 2)
                    Else
                        .ParameterFields(17).AddCurrentValue "Taxa FRC: " & FormatCurrency(totFRC + frcAponta + frcIntima + frcDistrib, 2)
                        totFRC = totFRC + frcAponta + frcIntima + frcDistrib
                    End If
                    DB.BeginTrans
                    DB.Execute ("Update tblFinanceiro set FRJ='" & Replace(totFRJ, ",", ".") & "',FRC='" & Replace(totFRC, ",", ".") & "' Where idFinanceiro='" & RSImp!idFinanceiro & "'")
                    DB.CommitTrans
                    If RSImp!Valor_Selo > 0 Then
                        .ParameterFields(18).AddCurrentValue "Selo Judicial: " & FormatCurrency(RSImp!Valor_Selo, 2)
                    End If
                    If RSImp!Valor_Distrib > 0 Then
                        .ParameterFields(19).AddCurrentValue "Distribuidor: " & FormatCurrency(eDistrib, 2)
                    End If
                    
                    .ParameterFields(20).AddCurrentValue "Desp. Apontamento: " & FormatCurrency(eAponta, 2)
                    .ParameterFields(21).AddCurrentValue "Desp. Intimação: " & FormatCurrency(eIntima + (eIntima * nAvalista), 2)
                    .ParameterFields(22).AddCurrentValue "Desp. Canc. Apont.: " & FormatCurrency(eCanc + eRet, 2)
                    .ParameterFields(23).AddCurrentValue "Desp. CPD: " & FormatCurrency(eCpd, 2)
                
                    If RS!Editado = True Then
                        .ParameterFields(24).AddCurrentValue "Edital: " & FormatCurrency(eEdital, 2)
                    End If
                    
                    '.ParameterFields(24).AddCurrentValue "TOTAL: " & FormatCurrency(RSImp!Pagar, 2)
                    .ParameterFields(25).AddCurrentValue "Adiantamento: " & FormatCurrency(Adiantamento, 2)

                    .ParameterFields(26).AddCurrentValue "TOTAL PAGO: " & FormatCurrency(RSImp!Pagar - Adiantamento, 2)
                    Data_Ocorrencia = Format(DataLiquida, "ddmmyyyy")
                    .ParameterFields(27).AddCurrentValue "" & "Belém, " & Day(CDate(Format(Data_Ocorrencia, "00/00/0000"))) & " de " & MonthName(Month(CDate(Format(Data_Ocorrencia, "00/00/0000")))) & " de " & Year(CDate(Format(Data_Ocorrencia, "00/00/0000")))
                    .ParameterFields(28).AddCurrentValue "Protocolo Dist.: " & RS!Protocolo_Dist
                    .ParameterFields(29).AddCurrentValue "ISS: " & FormatCurrency(RSImp!ISS, 2)
                End With
                Set CRExportOptions = CrysRep.ExportOptions
                CRExportOptions.FormatType = crEFTPortableDocFormat
                CRExportOptions.DestinationType = crEDTDiskFile
                CRExportOptions.DiskFileName = FileLoca
                CrysRep.DisplayProgressDialog = False
                CrysRep.Export False
                Set CRExportOptions = Nothing
            End If
            
            nResp = MsgBox("Imprimir o Recibo de Retirada?", vbQuestion + vbYesNo)
            If nResp = vbYes Then
                CrysRep.PrintOut False, 1
            End If

            Screen.MousePointer = 1
'            frmPrtCaixa.CRViewer1.ReportSource = CrysRep
'            frmPrtCaixa.CRViewer1.ViewReport
'            frmPrtCaixa.Show 1
            If Conn.State = 1 Then Conn.Close

            CaixaXML
'            EnvioXML

    '<<< Apaga QRCode >>>
            RSPrt.Open "Select * From tblCaixa", DB, adOpenDynamic
                Do While Not RSPrt.EOF
                    Kill RSPrt!QRCode
                    RSPrt.MoveNext
                Loop
            RSPrt.Close
            
            DB.Execute "Delete From tblCaixa"
'    RSBaixa.MoveNext
'    Loop

'<<< F I M   R E T I R A D A >>>
                
    Else
        MsgBox "O Título " & RSBaixa!Protocolo & " não pode ser RETIRADO.", vbInformation
    End If
        DB.Execute ("UPDATE tblGuias SET Baixado='" & 1 & "'WHERE N_Guia='" & RSBaixa!N_Guia & "'")
        RSBaixa.MoveNext
        
        If RS.State = 1 Then RS.Close
        If RSSeloDigital.State = 1 Then RSSeloDigital.Close
        If RSPostecipa.State = 1 Then RSPostecipa.Close
        If RSImp.State = 1 Then RSImp.Close
        If RSAdianta.State = 1 Then RSAdianta.Close
        
        Loop
        
    End If
        
    MeuGrid
    Carrega_Grid
    Me.optDataLiquida = False
    MsgBox "Liquidação finalizada com sucesso!", vbInformation
    
    Exit Sub
Erro:
    If Err.Number = 53 Then
        Resume Next
    End If
    MsgBox "Erro de Sistema " & Err.Number & Err.Description, vbCritical, "Erro"
'    Resume Next
    Close #1
'    DB.RollbackTrans
    Screen.MousePointer = 1

End Sub
Function Carrega_Grid()
On Error GoTo Erro
Dim Marca As String
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
Dim RSGuia As ADODB.Recordset
Set RSGuia = New ADODB.Recordset


If uCaixa = True Then
    Me.optDataLiquida.Visible = True
    RS.Open "SELECT N_Guia FROM tblGuias WHERE Marca='" & 1 & "'AND Baixado='" & 0 & "'GROUP BY N_Guia", DB, adOpenDynamic
    Me.txtTotalTitulos = RS.RecordCount
    If RS.RecordCount > 0 Then
        Do While Not RS.EOF
            RSGuia.Open "SELECT * FROM tblGuias WHERE N_Guia='" & RS!N_Guia & "'AND Baixado='" & 0 & "'ORDER BY N_Guia", DB, adOpenDynamic
            Pagar = 0
            Do While Not RSGuia.EOF
                If RSGuia!Marca = True Then
                    Marca = "Baixar"
                Else
                    Marca = ""
                End If
                N_Guia = RSGuia!N_Guia
                Pagar = Pagar + RSGuia!Pagar
                Tipo_Baixa = RSGuia!Tipo_Baixa
                Usuario = RSGuia!Usuario
                DataEntrada = RSGuia!DataEntrada
                Devedor = RSGuia!Devedor
                RSGuia.MoveNext
            Loop
            Me.GridGuias.AddItem 0 & vbTab & 0 & vbTab & Devedor & vbTab & Tipo_Baixa & vbTab & FormatCurrency(Pagar, 2) & vbTab & Usuario & vbTab & N_Guia & vbTab & Marca & vbTab & DataEntrada
'            Me.GridGuias.AddItem RS("Protocolo") & vbTab & RS("Ocorrencia") & vbTab & RS("Devedor") & vbTab & RS("Tipo_Baixa") & vbTab & FormatCurrency(RS("Pagar"), 2) & vbTab & RS("Usuario") & vbTab & RS("N_Guia") & vbTab & Marca & vbTab & RS("DataEntrada") & vbTab & Baixado
            Me.Refresh
            RSGuia.Close
        RS.MoveNext
        Loop
    End If
    Exit Function
'Else
'    RS.Open "SELECT * FROM tblGuias WHERE Baixado='" & 0 & "'ORDER BY N_Guia", DB, adOpenDynamic
End If

If sProtocolo <> "" Then
    'RS.Open "SELECT * FROM tblGuias WHERE Protocolo='" & sProtocolo & "'ORDER BY N_Guia", DB, adOpenDynamic
    RS.Open "SELECT N_Guia FROM tblGuias WHERE Protocolo='" & sProtocolo & "'GROUP BY N_Guia", DB, adOpenDynamic
End If
If sGuia <> "" Then
    RS.Open "SELECT N_Guia FROM tblGuias WHERE N_Guia='" & sGuia & "'GROUP BY N_Guia", DB, adOpenDynamic
End If
If sGuia = "" And sProtocolo = "" Then
    RS.Open "SELECT N_Guia FROM tblGuias WHERE DataEntrada BETWEEN'" & Format(sDataInicial, "mm/dd/yyyy") & "'AND'" & Format(sDataFinal, "mm/dd/yyyy") & "'GROUP BY N_Guia ORDER BY N_Guia DESC", DB, adOpenDynamic
End If
If RS.RecordCount = 0 Then
'    MsgBox "Dados inexistentes.", vbInformation
Else
    
    Do While Not RS.EOF
            RSGuia.Open "SELECT * FROM tblGuias WHERE N_Guia='" & RS!N_Guia & "'ORDER BY N_Guia DESC", DB, adOpenDynamic
        
            Pagar = 0
            Me.txtTotalTitulos = RSGuia.RecordCount
            Salva = 0
            Do While Not RSGuia.EOF
                If RSGuia!Marca = True Then
                    Marca = "Baixar"
                Else
                    Marca = ""
                End If

                If RSGuia!Baixado = True Then
                    Baixado = "SIM"
                Else
                    Baixado = "NÃO"
                End If
            
                N_Guia = RSGuia!N_Guia
                Pagar = Pagar + RSGuia!Pagar
                Tipo_Baixa = RSGuia!Tipo_Baixa
                Usuario = RSGuia!Usuario
                DataEntrada = RSGuia!DataEntrada
                Devedor = RSGuia!Devedor
                RSGuia.MoveNext
                Salva = 1
            Loop
            If Salva = 1 Then
                Me.GridGuias.AddItem 0 & vbTab & 0 & vbTab & Devedor & vbTab & Tipo_Baixa & vbTab & FormatCurrency(Pagar, 2) & vbTab & Usuario & vbTab & N_Guia & vbTab & Marca & vbTab & DataEntrada & vbTab & Baixado
            End If
        RSGuia.Close
'        Me.GridGuias.AddItem RS("Protocolo") & vbTab & RS("Ocorrencia") & vbTab & RS("Devedor") & vbTab & RS("Tipo_Baixa") & vbTab & FormatCurrency(RS("Pagar"), 2) & vbTab & RS("Usuario") & vbTab & RS("N_Guia") & vbTab & Marca & vbTab & RS("DataEntrada") & vbTab & Baixado
        Me.Refresh
    RS.MoveNext
    Loop
End If
ZEBRAR
'<<< Scroll Mouse >>>
'    WheelUnHook
'    WheelHook Me, GridGuias
'<<< Fim Carrega Títulos para protesto >>>
Me.cmdPesquisar.SetFocus
Exit Function
Erro:
    MsgBox "Erro de Sistema. " & Err.Description & " - N° " & Err.Number, vbCritical

End Function

Public Sub MeuGridTitulos()
    GridTitulos.Clear
    GridTitulos.FormatString = "Protocolo |Ocorrência |Devedor |Tipo Baixa |Total Pagar |Usuário |Nº Guia |Situação |Data |Liquidado "
    GridTitulos.Cols = 10
    GridTitulos.Rows = 1
    GridTitulos.FixedCols = 0
    GridTitulos.ColWidth(0) = 1000
    GridTitulos.ColWidth(1) = 1800
    GridTitulos.ColWidth(2) = 4000
    GridTitulos.ColWidth(3) = 1500
    GridTitulos.ColWidth(4) = 1100
    GridTitulos.ColWidth(5) = 1300
    GridTitulos.ColWidth(6) = 1000
    GridTitulos.ColWidth(7) = 1100
    GridTitulos.ColWidth(8) = 1000
    GridTitulos.ColWidth(9) = 800
    GridTitulos.ColAlignment(1) = 3
    GridTitulos.ColAlignment(3) = 3
    GridTitulos.ColAlignment(4) = 7
    GridTitulos.CellAlignment = 1
End Sub

Function Carrega_GridTitulos()
On Error GoTo Erro
Dim Marca As String
Dim RS As ADODB.Recordset
Set RS = New ADODB.Recordset
Dim RSGuia As ADODB.Recordset
Set RSGuia = New ADODB.Recordset

RS.Open "SELECT * FROM tblGuiaProvisoria WHERE Usuario='" & User & "'ORDER BY N_Guia", DB, adOpenDynamic
    Me.txtTotalTitulos = RS.RecordCount
    Me.txtPagar = 0
    Do While Not RS.EOF
        If RS!Marca = True Then
            Marca = "Baixar"
        Else
            Marca = ""
        End If
    
        If RS!Baixado = True Then
            Baixado = "SIM"
        Else
            Baixado = "NÃO"
        End If
        Me.txtPagar = FormatCurrency(CDbl(txtPagar) + RS!Pagar, 2)
        Me.GridTitulos.AddItem RS("Protocolo") & vbTab & RS("Ocorrencia") & vbTab & RS("Devedor") & vbTab & RS("Tipo_Baixa") & vbTab & FormatCurrency(RS("Pagar"), 2) & vbTab & RS("Usuario") & vbTab & RS("N_Guia") & vbTab & Marca & vbTab & RS("DataEntrada") & vbTab & Baixado
        Me.Refresh
    RS.MoveNext
    Loop

Me.cmdPesquisar.SetFocus
Me.Refresh
Exit Function
Erro:
    MsgBox "Erro de Sistema. " & Err.Description & " - N° " & Err.Number, vbCritical

End Function

Function EIMPAR(ByVal INUM As Long) As Boolean
EIMPAR = (INUM Mod 2)
End Function

Sub FLEXCORES(LCORPAR As Long, LCORIMPAR As Long)
Dim ILINHA As Integer
GridGuias.FillStyle = flexFillRepeat
For ILINHA = 1 To GridGuias.Rows - 1
With GridGuias
.Row = ILINHA
If EIMPAR(ILINHA) Then
.Col = 0
.ColSel = .Cols - 1
.CellBackColor = LCORIMPAR
Else
.Col = 1
.ColSel = .Cols - 1
.CellBackColor = LCORPAR
End If
End With
Next
GridGuias.FillStyle = flexFillSingle
End Sub

Private Sub ZEBRAR()
FLEXCORES (&HFFFFFF), (&HC0FFFF)
End Sub

Public Sub CalculoFaixas(Aponta As ADODB.Recordset)
If Tipo_Baixa = "Pagamento" Or Tipo_Baixa = "Retirada" Then
        If Valor <= Aponta!Faixa0 Then txtValorCustas = FormatCurrency(Aponta!Pago0, 2): vPago = Aponta!sPago0: atoPago = 761: nFaixa0 = 0: atoCanc = 894
        If Valor <= Aponta!Faixa1 And Valor > Aponta!Faixa0 Then txtValorCustas = FormatCurrency(Aponta!Pago1, 2): vPago = Aponta!sPago1: atoPago = 762: atoCanc = 895
        If Valor <= Aponta!Faixa2 And Valor > Aponta!Faixa1 Then txtValorCustas = FormatCurrency(Aponta!Pago2, 2): vPago = Aponta!sPago2: atoPago = 763: atoCanc = 896
        If Valor <= Aponta!Faixa3 And Valor > Aponta!Faixa2 Then txtValorCustas = FormatCurrency(Aponta!Pago3, 2): vPago = Aponta!sPago3: atoPago = 764: atoCanc = 897
        If Valor <= Aponta!Faixa4 And Valor > Aponta!Faixa3 Then txtValorCustas = FormatCurrency(Aponta!Pago4, 2): vPago = Aponta!sPago4: atoPago = 765: atoCanc = 898
        If Valor <= Aponta!Faixa5 And Valor > Aponta!Faixa4 Then txtValorCustas = FormatCurrency(Aponta!Pago5, 2): vPago = Aponta!sPago5: atoPago = 766: atoCanc = 899
        If Valor <= Aponta!Faixa6 And Valor > Aponta!Faixa5 Then txtValorCustas = FormatCurrency(Aponta!Pago6, 2): vPago = Aponta!sPago6: atoPago = 767: atoCanc = 900
        If Valor <= Aponta!Faixa7 And Valor > Aponta!Faixa6 Then txtValorCustas = FormatCurrency(Aponta!Pago7, 2): vPago = Aponta!sPago7: atoPago = 768: atoCanc = 901
        If Valor <= Aponta!Faixa8 And Valor > Aponta!Faixa7 Then txtValorCustas = FormatCurrency(Aponta!Pago8, 2): vPago = Aponta!sPago8: atoPago = 769: atoCanc = 902
        If Valor <= Aponta!Faixa9 And Valor > Aponta!Faixa8 Then txtValorCustas = FormatCurrency(Aponta!Pago9, 2): vPago = Aponta!sPago9: atoPago = 770: atoCanc = 903
        If Valor <= Aponta!Faixa10 And Valor > Aponta!Faixa9 Then txtValorCustas = FormatCurrency(Aponta!Pago10, 2): vPago = Aponta!sPago10: atoPago = 771: atoCanc = 904
        If Valor <= Aponta!Faixa11 And Valor > Aponta!Faixa10 Then txtValorCustas = FormatCurrency(Aponta!Pago11, 2): vPago = Aponta!sPago11: atoPago = 772: atoCanc = 905
        If Valor <= Aponta!Faixa12 And Valor > Aponta!Faixa11 Then txtValorCustas = FormatCurrency(Aponta!Pago12, 2): vPago = Aponta!sPago12: atoPago = 773: atoCanc = 906
        If Valor <= Aponta!Faixa13 And Valor > Aponta!Faixa12 Then txtValorCustas = FormatCurrency(Aponta!Pago13, 2): vPago = Aponta!sPago13: atoPago = 774: atoCanc = 907
        If Valor <= Aponta!Faixa14 And Valor > Aponta!Faixa13 Then txtValorCustas = FormatCurrency(Aponta!Pago14, 2): vPago = Aponta!sPago14: atoPago = 775: atoCanc = 908
        If Valor <= Aponta!Faixa15 And Valor > Aponta!Faixa14 Then txtValorCustas = FormatCurrency(Aponta!Pago15, 2): vPago = Aponta!sPago15: atoPago = 776: atoCanc = 909
        If Valor <= Aponta!Faixa16 And Valor > Aponta!Faixa15 Then txtValorCustas = FormatCurrency(Aponta!Pago16, 2): vPago = Aponta!sPago16: atoPago = 777: atoCanc = 910
        If Valor <= Aponta!Faixa17 And Valor > Aponta!Faixa16 Then txtValorCustas = FormatCurrency(Aponta!Pago17, 2): vPago = Aponta!sPago17: atoPago = 778: atoCanc = 911
        If Valor <= Aponta!Faixa18 And Valor > Aponta!Faixa17 Then txtValorCustas = FormatCurrency(Aponta!Pago18, 2): vPago = Aponta!sPago18: atoPago = 779: atoCanc = 912
        If Valor <= Aponta!Faixa19 And Valor > Aponta!Faixa18 Then txtValorCustas = FormatCurrency(Aponta!Pago19, 2): vPago = Aponta!sPago19: atoPago = 780: atoCanc = 913
        If Valor <= Aponta!Faixa20 And Valor > Aponta!Faixa19 Then txtValorCustas = FormatCurrency(Aponta!Pago20, 2): vPago = Aponta!sPago20: atoPago = 781: atoCanc = 914
        If Valor <= Aponta!Faixa21 And Valor > Aponta!Faixa20 Then txtValorCustas = FormatCurrency(Aponta!Pago21, 2): vPago = Aponta!sPago21: atoPago = 782: atoCanc = 915
        If Valor <= Aponta!Faixa22 And Valor > Aponta!Faixa21 Then txtValorCustas = FormatCurrency(Aponta!Pago22, 2): vPago = Aponta!sPago22: atoPago = 783: atoCanc = 916
        If Valor <= Aponta!Faixa23 And Valor > Aponta!Faixa22 Then txtValorCustas = FormatCurrency(Aponta!Pago23, 2): vPago = Aponta!sPago23: atoPago = 784: atoCanc = 917
        If Valor <= Aponta!Faixa24 And Valor > Aponta!Faixa23 Then txtValorCustas = FormatCurrency(Aponta!Pago24, 2): vPago = Aponta!sPago24: atoPago = 785: atoCanc = 918
        If Valor <= Aponta!Faixa25 And Valor > Aponta!Faixa24 Then txtValorCustas = FormatCurrency(Aponta!Pago25, 2): vPago = Aponta!sPago25: atoPago = 786: atoCanc = 919
        If Valor <= Aponta!Faixa26 And Valor > Aponta!Faixa25 Then txtValorCustas = FormatCurrency(Aponta!Pago26, 2): vPago = Aponta!sPago26: atoPago = 787: atoCanc = 920
        If Valor <= Aponta!Faixa27 And Valor > Aponta!Faixa26 Then txtValorCustas = FormatCurrency(Aponta!Pago27, 2): vPago = Aponta!sPago27: atoPago = 788: atoCanc = 921
        If Valor <= Aponta!Faixa28 And Valor > Aponta!Faixa27 Then txtValorCustas = FormatCurrency(Aponta!Pago28, 2): vPago = Aponta!sPago28: atoPago = 789: atoCanc = 922
        If Valor <= Aponta!Faixa29 And Valor > Aponta!Faixa28 Then txtValorCustas = FormatCurrency(Aponta!Pago29, 2): vPago = Aponta!sPago29: atoPago = 790: atoCanc = 923
        If Valor <= Aponta!Faixa30 And Valor > Aponta!Faixa29 Then txtValorCustas = FormatCurrency(Aponta!Pago30, 2): vPago = Aponta!sPago30: atoPago = 791: atoCanc = 924
        If Valor <= Aponta!Faixa31 And Valor > Aponta!Faixa30 Then txtValorCustas = FormatCurrency(Aponta!Pago31, 2): vPago = Aponta!sPago31: atoPago = 792: atoCanc = 925
        If Valor <= Aponta!Faixa32 And Valor > Aponta!Faixa31 Then txtValorCustas = FormatCurrency(Aponta!Pago32, 2): vPago = Aponta!sPago32: atoPago = 793: atoCanc = 926
        If Valor <= Aponta!Faixa33 And Valor > Aponta!Faixa32 Then txtValorCustas = FormatCurrency(Aponta!Pago33, 2): vPago = Aponta!sPago33: atoPago = 794: atoCanc = 927
        If Valor <= Aponta!Faixa34 And Valor > Aponta!Faixa33 Then txtValorCustas = FormatCurrency(Aponta!Pago34, 2): vPago = Aponta!sPago34: atoPago = 795: atoCanc = 928
        If Valor <= Aponta!Faixa35 And Valor > Aponta!Faixa34 Then txtValorCustas = FormatCurrency(Aponta!Pago35, 2): vPago = Aponta!sPago35: atoPago = 796: atoCanc = 929
        If Valor <= Aponta!Faixa36 And Valor > Aponta!Faixa35 Then txtValorCustas = FormatCurrency(Aponta!Pago36, 2): vPago = Aponta!sPago36: atoPago = 797: atoCanc = 930
        If Valor <= Aponta!Faixa37 And Valor > Aponta!Faixa36 Then txtValorCustas = FormatCurrency(Aponta!Pago37, 2): vPago = Aponta!sPago37: atoPago = 798: atoCanc = 931
        If Valor <= Aponta!Faixa38 And Valor > Aponta!Faixa37 Then txtValorCustas = FormatCurrency(Aponta!Pago38, 2): vPago = Aponta!sPago38: atoPago = 799: atoCanc = 932
        If Valor <= Aponta!Faixa39 And Valor > Aponta!Faixa38 Then txtValorCustas = FormatCurrency(Aponta!Pago39, 2): vPago = Aponta!sPago39: atoPago = 800: atoCanc = 933
        If Valor <= Aponta!Faixa40 And Valor > Aponta!Faixa39 Then txtValorCustas = FormatCurrency(Aponta!Pago40, 2): vPago = Aponta!sPago40: atoPago = 801: atoCanc = 934
        If Valor <= Aponta!Faixa41 And Valor > Aponta!Faixa40 Then txtValorCustas = FormatCurrency(Aponta!Pago41, 2): vPago = Aponta!sPago41: atoPago = 802: atoCanc = 935
        If Valor <= Aponta!Faixa42 And Valor > Aponta!Faixa41 Then txtValorCustas = FormatCurrency(Aponta!Pago42, 2): vPago = Aponta!sPago42: atoPago = 803: atoCanc = 936
        If Valor <= Aponta!Faixa43 And Valor > Aponta!Faixa42 Then txtValorCustas = FormatCurrency(Aponta!Pago43, 2): vPago = Aponta!sPago43: atoPago = 804: atoCanc = 937
        If Valor <= Aponta!Faixa44 And Valor > Aponta!Faixa43 Then txtValorCustas = FormatCurrency(Aponta!Pago44, 2): vPago = Aponta!sPago44: atoPago = 805: atoCanc = 938
        If Valor <= Aponta!Faixa45 And Valor > Aponta!Faixa44 Then txtValorCustas = FormatCurrency(Aponta!Pago45, 2): vPago = Aponta!sPago45: atoPago = 806: atoCanc = 939
        If Valor <= Aponta!Faixa46 And Valor > Aponta!Faixa45 Then txtValorCustas = FormatCurrency(Aponta!Pago46, 2): vPago = Aponta!sPago46: atoPago = 807: atoCanc = 940
        If Valor <= Aponta!Faixa47 And Valor > Aponta!Faixa46 Then txtValorCustas = FormatCurrency(Aponta!Pago47, 2): vPago = Aponta!sPago47: atoPago = 808: atoCanc = 941
        If Valor <= Aponta!Faixa48 And Valor > Aponta!Faixa47 Then txtValorCustas = FormatCurrency(Aponta!Pago48, 2): vPago = Aponta!sPago48: atoPago = 809: atoCanc = 942
        If Valor <= Aponta!Faixa49 And Valor > Aponta!Faixa48 Then txtValorCustas = FormatCurrency(Aponta!Pago49, 2): vPago = Aponta!sPago49: atoPago = 810: atoCanc = 943
        If Valor <= Aponta!Faixa50 And Valor > Aponta!Faixa49 Then txtValorCustas = FormatCurrency(Aponta!Pago50, 2): vPago = Aponta!sPago50: atoPago = 811: atoCanc = 944
        If Valor <= Aponta!Faixa51 And Valor > Aponta!Faixa50 Then txtValorCustas = FormatCurrency(Aponta!Pago51, 2): vPago = Aponta!sPago51: atoPago = 812: atoCanc = 945
        If Valor <= Aponta!Faixa52 And Valor > Aponta!Faixa51 Then txtValorCustas = FormatCurrency(Aponta!Pago52, 2): vPago = Aponta!sPago52: atoPago = 813: atoCanc = 946
        If Valor <= Aponta!Faixa53 And Valor > Aponta!Faixa52 Then txtValorCustas = FormatCurrency(Aponta!Pago53, 2): vPago = Aponta!sPago53: atoPago = 814: atoCanc = 947
        If Valor <= Aponta!Faixa54 And Valor > Aponta!Faixa53 Then txtValorCustas = FormatCurrency(Aponta!Pago54, 2): vPago = Aponta!sPago54: atoPago = 815: atoCanc = 948
        If Valor <= Aponta!Faixa55 And Valor > Aponta!Faixa54 Then txtValorCustas = FormatCurrency(Aponta!Pago55, 2): vPago = Aponta!sPago55: atoPago = 816: atoCanc = 949
        If Valor <= Aponta!Faixa56 And Valor > Aponta!Faixa55 Then txtValorCustas = FormatCurrency(Aponta!Pago56, 2): vPago = Aponta!sPago56: atoPago = 817: atoCanc = 950
        If Valor <= Aponta!Faixa57 And Valor > Aponta!Faixa56 Then txtValorCustas = FormatCurrency(Aponta!Pago57, 2): vPago = Aponta!sPago57: atoPago = 818: atoCanc = 951
        If Valor <= Aponta!Faixa58 And Valor > Aponta!Faixa57 Then txtValorCustas = FormatCurrency(Aponta!Pago58, 2): vPago = Aponta!sPago58: atoPago = 819: atoCanc = 952
        If Valor <= Aponta!Faixa59 And Valor > Aponta!Faixa58 Then txtValorCustas = FormatCurrency(Aponta!Pago59, 2): vPago = Aponta!sPago59: atoPago = 820: atoCanc = 953
        If Valor <= Aponta!Faixa60 And Valor > Aponta!Faixa59 Then txtValorCustas = FormatCurrency(Aponta!Pago60, 2): vPago = Aponta!sPago60: atoPago = 821: atoCanc = 954
        If Valor <= Aponta!Faixa61 And Valor > Aponta!Faixa60 Then txtValorCustas = FormatCurrency(Aponta!Pago61, 2): vPago = Aponta!sPago61: atoPago = 822: atoCanc = 955
        If Valor <= Aponta!Faixa62 And Valor > Aponta!Faixa61 Then txtValorCustas = FormatCurrency(Aponta!Pago62, 2): vPago = Aponta!sPago62: atoPago = 823: atoCanc = 956
        If Valor <= Aponta!Faixa63 And Valor > Aponta!Faixa62 Then txtValorCustas = FormatCurrency(Aponta!Pago63, 2): vPago = Aponta!sPago63: atoPago = 824: atoCanc = 957
        If Valor <= Aponta!Faixa64 And Valor > Aponta!Faixa63 Then txtValorCustas = FormatCurrency(Aponta!Pago64, 2): vPago = Aponta!sPago64: atoPago = 825: atoCanc = 958
        If Valor > Aponta!Faixa64 Then txtValorCustas = FormatCurrency(Aponta!Pago65, 2): vPago = Aponta!sPago65: atoPago = 826: atoCanc = 959
        vRet = vPago

End If
        
If Tipo_Baixa = "CANCELAMENTO" Or Tipo_Baixa = "Cancelamento" Then
        If Valor <= Aponta!Faixa0 Then txtValorCustas = FormatCurrency(Aponta!Cancelado0, 2): vPago = Aponta!sPago0: atoPago = 761: nFaixa0 = 0: atoCanc = 894
        If Valor <= Aponta!Faixa1 And Valor > Aponta!Faixa0 Then txtValorCustas = FormatCurrency(Aponta!Cancelado1, 2): vPago = Aponta!sPago1: atoPago = 762: atoCanc = 895
        If Valor <= Aponta!Faixa2 And Valor > Aponta!Faixa1 Then txtValorCustas = FormatCurrency(Aponta!Cancelado2, 2): vPago = Aponta!sPago2: atoPago = 763: atoCanc = 896
        If Valor <= Aponta!Faixa3 And Valor > Aponta!Faixa2 Then txtValorCustas = FormatCurrency(Aponta!Cancelado3, 2): vPago = Aponta!sPago3: atoPago = 764: atoCanc = 897
        If Valor <= Aponta!Faixa4 And Valor > Aponta!Faixa3 Then txtValorCustas = FormatCurrency(Aponta!Cancelado4, 2): vPago = Aponta!sPago4: atoPago = 765: atoCanc = 898
        If Valor <= Aponta!Faixa5 And Valor > Aponta!Faixa4 Then txtValorCustas = FormatCurrency(Aponta!Cancelado5, 2): vPago = Aponta!sPago5: atoPago = 766: atoCanc = 899
        If Valor <= Aponta!Faixa6 And Valor > Aponta!Faixa5 Then txtValorCustas = FormatCurrency(Aponta!Cancelado6, 2): vPago = Aponta!sPago6: atoPago = 767: atoCanc = 900
        If Valor <= Aponta!Faixa7 And Valor > Aponta!Faixa6 Then txtValorCustas = FormatCurrency(Aponta!Cancelado7, 2): vPago = Aponta!sPago7: atoPago = 768: atoCanc = 901
        If Valor <= Aponta!Faixa8 And Valor > Aponta!Faixa7 Then txtValorCustas = FormatCurrency(Aponta!Cancelado8, 2): vPago = Aponta!sPago8: atoPago = 769: atoCanc = 902
        If Valor <= Aponta!Faixa9 And Valor > Aponta!Faixa8 Then txtValorCustas = FormatCurrency(Aponta!Cancelado9, 2): vPago = Aponta!sPago9: atoPago = 770: atoCanc = 903
        If Valor <= Aponta!Faixa10 And Valor > Aponta!Faixa9 Then txtValorCustas = FormatCurrency(Aponta!Cancelado10, 2): vPago = Aponta!sPago10: atoPago = 771: atoCanc = 904
        If Valor <= Aponta!Faixa11 And Valor > Aponta!Faixa10 Then txtValorCustas = FormatCurrency(Aponta!Cancelado11, 2): vPago = Aponta!sPago11: atoPago = 772: atoCanc = 905
        If Valor <= Aponta!Faixa12 And Valor > Aponta!Faixa11 Then txtValorCustas = FormatCurrency(Aponta!Cancelado12, 2): vPago = Aponta!sPago12: atoPago = 773: atoCanc = 906
        If Valor <= Aponta!Faixa13 And Valor > Aponta!Faixa12 Then txtValorCustas = FormatCurrency(Aponta!Cancelado13, 2): vPago = Aponta!sPago13: atoPago = 774: atoCanc = 907
        If Valor <= Aponta!Faixa14 And Valor > Aponta!Faixa13 Then txtValorCustas = FormatCurrency(Aponta!Cancelado14, 2): vPago = Aponta!sPago14: atoPago = 775: atoCanc = 908
        If Valor <= Aponta!Faixa15 And Valor > Aponta!Faixa14 Then txtValorCustas = FormatCurrency(Aponta!Cancelado15, 2): vPago = Aponta!sPago15: atoPago = 776: atoCanc = 909
        If Valor <= Aponta!Faixa16 And Valor > Aponta!Faixa15 Then txtValorCustas = FormatCurrency(Aponta!Cancelado16, 2): vPago = Aponta!sPago16: atoPago = 777: atoCanc = 910
        If Valor <= Aponta!Faixa17 And Valor > Aponta!Faixa16 Then txtValorCustas = FormatCurrency(Aponta!Cancelado17, 2): vPago = Aponta!sPago17: atoPago = 778: atoCanc = 911
        If Valor <= Aponta!Faixa18 And Valor > Aponta!Faixa17 Then txtValorCustas = FormatCurrency(Aponta!Cancelado18, 2): vPago = Aponta!sPago18: atoPago = 779: atoCanc = 912
        If Valor <= Aponta!Faixa19 And Valor > Aponta!Faixa18 Then txtValorCustas = FormatCurrency(Aponta!Cancelado19, 2): vPago = Aponta!sPago19: atoPago = 780: atoCanc = 913
        If Valor <= Aponta!Faixa20 And Valor > Aponta!Faixa19 Then txtValorCustas = FormatCurrency(Aponta!Cancelado20, 2): vPago = Aponta!sPago20: atoPago = 781: atoCanc = 914
        If Valor <= Aponta!Faixa21 And Valor > Aponta!Faixa20 Then txtValorCustas = FormatCurrency(Aponta!Cancelado21, 2): vPago = Aponta!sPago21: atoPago = 782: atoCanc = 915
        If Valor <= Aponta!Faixa22 And Valor > Aponta!Faixa21 Then txtValorCustas = FormatCurrency(Aponta!Cancelado22, 2): vPago = Aponta!sPago22: atoPago = 783: atoCanc = 916
        If Valor <= Aponta!Faixa23 And Valor > Aponta!Faixa22 Then txtValorCustas = FormatCurrency(Aponta!Cancelado23, 2): vPago = Aponta!sPago23: atoPago = 784: atoCanc = 917
        If Valor <= Aponta!Faixa24 And Valor > Aponta!Faixa23 Then txtValorCustas = FormatCurrency(Aponta!Cancelado24, 2): vPago = Aponta!sPago24: atoPago = 785: atoCanc = 918
        If Valor <= Aponta!Faixa25 And Valor > Aponta!Faixa24 Then txtValorCustas = FormatCurrency(Aponta!Cancelado25, 2): vPago = Aponta!sPago25: atoPago = 786: atoCanc = 919
        If Valor <= Aponta!Faixa26 And Valor > Aponta!Faixa25 Then txtValorCustas = FormatCurrency(Aponta!Cancelado26, 2): vPago = Aponta!sPago26: atoPago = 787: atoCanc = 920
        If Valor <= Aponta!Faixa27 And Valor > Aponta!Faixa26 Then txtValorCustas = FormatCurrency(Aponta!Cancelado27, 2): vPago = Aponta!sPago27: atoPago = 788: atoCanc = 921
        If Valor <= Aponta!Faixa28 And Valor > Aponta!Faixa27 Then txtValorCustas = FormatCurrency(Aponta!Cancelado28, 2): vPago = Aponta!sPago28: atoPago = 789: atoCanc = 922
        If Valor <= Aponta!Faixa29 And Valor > Aponta!Faixa28 Then txtValorCustas = FormatCurrency(Aponta!Cancelado29, 2): vPago = Aponta!sPago29: atoPago = 790: atoCanc = 923
        If Valor <= Aponta!Faixa30 And Valor > Aponta!Faixa29 Then txtValorCustas = FormatCurrency(Aponta!Cancelado30, 2): vPago = Aponta!sPago30: atoPago = 791: atoCanc = 924
        If Valor <= Aponta!Faixa31 And Valor > Aponta!Faixa30 Then txtValorCustas = FormatCurrency(Aponta!Cancelado31, 2): vPago = Aponta!sPago31: atoPago = 792: atoCanc = 925
        If Valor <= Aponta!Faixa32 And Valor > Aponta!Faixa31 Then txtValorCustas = FormatCurrency(Aponta!Cancelado32, 2): vPago = Aponta!sPago32: atoPago = 793: atoCanc = 926
        If Valor <= Aponta!Faixa33 And Valor > Aponta!Faixa32 Then txtValorCustas = FormatCurrency(Aponta!Cancelado33, 2): vPago = Aponta!sPago33: atoPago = 794: atoCanc = 927
        If Valor <= Aponta!Faixa34 And Valor > Aponta!Faixa33 Then txtValorCustas = FormatCurrency(Aponta!Cancelado34, 2): vPago = Aponta!sPago34: atoPago = 795: atoCanc = 928
        If Valor <= Aponta!Faixa35 And Valor > Aponta!Faixa34 Then txtValorCustas = FormatCurrency(Aponta!Cancelado35, 2): vPago = Aponta!sPago35: atoPago = 796: atoCanc = 929
        If Valor <= Aponta!Faixa36 And Valor > Aponta!Faixa35 Then txtValorCustas = FormatCurrency(Aponta!Cancelado36, 2): vPago = Aponta!sPago36: atoPago = 797: atoCanc = 930
        If Valor <= Aponta!Faixa37 And Valor > Aponta!Faixa36 Then txtValorCustas = FormatCurrency(Aponta!Cancelado37, 2): vPago = Aponta!sPago37: atoPago = 798: atoCanc = 931
        If Valor <= Aponta!Faixa38 And Valor > Aponta!Faixa37 Then txtValorCustas = FormatCurrency(Aponta!Cancelado38, 2): vPago = Aponta!sPago38: atoPago = 799: atoCanc = 932
        If Valor <= Aponta!Faixa39 And Valor > Aponta!Faixa38 Then txtValorCustas = FormatCurrency(Aponta!Cancelado39, 2): vPago = Aponta!sPago39: atoPago = 800: atoCanc = 933
        If Valor <= Aponta!Faixa40 And Valor > Aponta!Faixa39 Then txtValorCustas = FormatCurrency(Aponta!Cancelado40, 2): vPago = Aponta!sPago40: atoPago = 801: atoCanc = 934
        If Valor <= Aponta!Faixa41 And Valor > Aponta!Faixa40 Then txtValorCustas = FormatCurrency(Aponta!Cancelado41, 2): vPago = Aponta!sPago41: atoPago = 802: atoCanc = 935
        If Valor <= Aponta!Faixa42 And Valor > Aponta!Faixa41 Then txtValorCustas = FormatCurrency(Aponta!Cancelado42, 2): vPago = Aponta!sPago42: atoPago = 803: atoCanc = 936
        If Valor <= Aponta!Faixa43 And Valor > Aponta!Faixa42 Then txtValorCustas = FormatCurrency(Aponta!Cancelado43, 2): vPago = Aponta!sPago43: atoPago = 804: atoCanc = 937
        If Valor <= Aponta!Faixa44 And Valor > Aponta!Faixa43 Then txtValorCustas = FormatCurrency(Aponta!Cancelado44, 2): vPago = Aponta!sPago44: atoPago = 805: atoCanc = 938
        If Valor <= Aponta!Faixa45 And Valor > Aponta!Faixa44 Then txtValorCustas = FormatCurrency(Aponta!Cancelado45, 2): vPago = Aponta!sPago45: atoPago = 806: atoCanc = 939
        If Valor <= Aponta!Faixa46 And Valor > Aponta!Faixa45 Then txtValorCustas = FormatCurrency(Aponta!Cancelado46, 2): vPago = Aponta!sPago46: atoPago = 807: atoCanc = 940
        If Valor <= Aponta!Faixa47 And Valor > Aponta!Faixa46 Then txtValorCustas = FormatCurrency(Aponta!Cancelado47, 2): vPago = Aponta!sPago47: atoPago = 808: atoCanc = 941
        If Valor <= Aponta!Faixa48 And Valor > Aponta!Faixa47 Then txtValorCustas = FormatCurrency(Aponta!Cancelado48, 2): vPago = Aponta!sPago48: atoPago = 809: atoCanc = 942
        If Valor <= Aponta!Faixa49 And Valor > Aponta!Faixa48 Then txtValorCustas = FormatCurrency(Aponta!Cancelado49, 2): vPago = Aponta!sPago49: atoPago = 810: atoCanc = 943
        If Valor <= Aponta!Faixa50 And Valor > Aponta!Faixa49 Then txtValorCustas = FormatCurrency(Aponta!Cancelado50, 2): vPago = Aponta!sPago50: atoPago = 811: atoCanc = 944
        If Valor <= Aponta!Faixa51 And Valor > Aponta!Faixa50 Then txtValorCustas = FormatCurrency(Aponta!Cancelado51, 2): vPago = Aponta!sPago51: atoPago = 812: atoCanc = 945
        If Valor <= Aponta!Faixa52 And Valor > Aponta!Faixa51 Then txtValorCustas = FormatCurrency(Aponta!Cancelado52, 2): vPago = Aponta!sPago52: atoPago = 813: atoCanc = 946
        If Valor <= Aponta!Faixa53 And Valor > Aponta!Faixa52 Then txtValorCustas = FormatCurrency(Aponta!Cancelado53, 2): vPago = Aponta!sPago53: atoPago = 814: atoCanc = 947
        If Valor <= Aponta!Faixa54 And Valor > Aponta!Faixa53 Then txtValorCustas = FormatCurrency(Aponta!Cancelado54, 2): vPago = Aponta!sPago54: atoPago = 815: atoCanc = 948
        If Valor <= Aponta!Faixa55 And Valor > Aponta!Faixa54 Then txtValorCustas = FormatCurrency(Aponta!Cancelado55, 2): vPago = Aponta!sPago55: atoPago = 816: atoCanc = 949
        If Valor <= Aponta!Faixa56 And Valor > Aponta!Faixa55 Then txtValorCustas = FormatCurrency(Aponta!Cancelado56, 2): vPago = Aponta!sPago56: atoPago = 817: atoCanc = 950
        If Valor <= Aponta!Faixa57 And Valor > Aponta!Faixa56 Then txtValorCustas = FormatCurrency(Aponta!Cancelado57, 2): vPago = Aponta!sPago57: atoPago = 818: atoCanc = 951
        If Valor <= Aponta!Faixa58 And Valor > Aponta!Faixa57 Then txtValorCustas = FormatCurrency(Aponta!Cancelado58, 2): vPago = Aponta!sPago58: atoPago = 819: atoCanc = 952
        If Valor <= Aponta!Faixa59 And Valor > Aponta!Faixa58 Then txtValorCustas = FormatCurrency(Aponta!Cancelado59, 2): vPago = Aponta!sPago59: atoPago = 820: atoCanc = 953
        If Valor <= Aponta!Faixa60 And Valor > Aponta!Faixa59 Then txtValorCustas = FormatCurrency(Aponta!Cancelado60, 2): vPago = Aponta!sPago60: atoPago = 821: atoCanc = 954
        If Valor <= Aponta!Faixa61 And Valor > Aponta!Faixa60 Then txtValorCustas = FormatCurrency(Aponta!Cancelado61, 2): vPago = Aponta!sPago61: atoPago = 822: atoCanc = 955
        If Valor <= Aponta!Faixa62 And Valor > Aponta!Faixa61 Then txtValorCustas = FormatCurrency(Aponta!Cancelado62, 2): vPago = Aponta!sPago62: atoPago = 823: atoCanc = 956
        If Valor <= Aponta!Faixa63 And Valor > Aponta!Faixa62 Then txtValorCustas = FormatCurrency(Aponta!Cancelado63, 2): vPago = Aponta!sPago63: atoPago = 824: atoCanc = 957
        If Valor <= Aponta!Faixa64 And Valor > Aponta!Faixa63 Then txtValorCustas = FormatCurrency(Aponta!Cancelado64, 2): vPago = Aponta!sPago64: atoPago = 825: atoCanc = 958
        If Valor > Aponta!Faixa64 Then txtValorCustas = FormatCurrency(Aponta!Cancelado65, 2): vPago = Aponta!sPago65: atoPago = 826: atoCanc = 959
End If
End Sub

Public Sub CalculoFaixaProtesto(Aponta As ADODB.Recordset)
        If Valor <= Aponta!Faixa0 Then txtValorCustas = FormatCurrency(Aponta!Pago720, 2): vPago = Aponta!sPago0: atoPago = 761: nFaixa0 = 0: atoCanc = 894: vProt = Aponta!sProt0: atoProt = 827
        If Valor <= Aponta!Faixa1 And Valor > Aponta!Faixa0 Then txtValorCustas = FormatCurrency(Aponta!Pago721, 2): vPago = Aponta!sPago1: atoPago = 762: atoCanc = 895: vProt = Aponta!sProt1: atoProt = 828
        If Valor <= Aponta!Faixa2 And Valor > Aponta!Faixa1 Then txtValorCustas = FormatCurrency(Aponta!Pago722, 2): vPago = Aponta!sPago2: atoPago = 763: atoCanc = 896: vProt = Aponta!sProt2: atoProt = 829
        If Valor <= Aponta!Faixa3 And Valor > Aponta!Faixa2 Then txtValorCustas = FormatCurrency(Aponta!Pago723, 2): vPago = Aponta!sPago3: atoPago = 764: atoCanc = 897: vProt = Aponta!sProt3: atoProt = 830
        If Valor <= Aponta!Faixa4 And Valor > Aponta!Faixa3 Then txtValorCustas = FormatCurrency(Aponta!Pago724, 2): vPago = Aponta!sPago4: atoPago = 765: atoCanc = 898: vProt = Aponta!sProt4: atoProt = 831
        If Valor <= Aponta!Faixa5 And Valor > Aponta!Faixa4 Then txtValorCustas = FormatCurrency(Aponta!Pago725, 2): vPago = Aponta!sPago5: atoPago = 766: atoCanc = 899: vProt = Aponta!sProt5: atoProt = 832
        If Valor <= Aponta!Faixa6 And Valor > Aponta!Faixa5 Then txtValorCustas = FormatCurrency(Aponta!Pago726, 2): vPago = Aponta!sPago6: atoPago = 767: atoCanc = 900: vProt = Aponta!sProt6: atoProt = 833
        If Valor <= Aponta!Faixa7 And Valor > Aponta!Faixa6 Then txtValorCustas = FormatCurrency(Aponta!Pago727, 2): vPago = Aponta!sPago7: atoPago = 768: atoCanc = 901: vProt = Aponta!sProt7: atoProt = 834
        If Valor <= Aponta!Faixa8 And Valor > Aponta!Faixa7 Then txtValorCustas = FormatCurrency(Aponta!Pago728, 2): vPago = Aponta!sPago8: atoPago = 769: atoCanc = 902: vProt = Aponta!sProt8: atoProt = 835
        If Valor <= Aponta!Faixa9 And Valor > Aponta!Faixa8 Then txtValorCustas = FormatCurrency(Aponta!Pago729, 2): vPago = Aponta!sPago9: atoPago = 770: atoCanc = 903: vProt = Aponta!sProt9: atoProt = 836
        If Valor <= Aponta!Faixa10 And Valor > Aponta!Faixa9 Then txtValorCustas = FormatCurrency(Aponta!Pago7210, 2): vPago = Aponta!sPago10: atoPago = 771: atoCanc = 904: vProt = Aponta!sProt10: atoProt = 837
        If Valor <= Aponta!Faixa11 And Valor > Aponta!Faixa10 Then txtValorCustas = FormatCurrency(Aponta!Pago7211, 2): vPago = Aponta!sPago11: atoPago = 772: atoCanc = 905: vProt = Aponta!sProt11: atoProt = 838
        If Valor <= Aponta!Faixa12 And Valor > Aponta!Faixa11 Then txtValorCustas = FormatCurrency(Aponta!Pago7212, 2): vPago = Aponta!sPago12: atoPago = 773: atoCanc = 906: vProt = Aponta!sProt12: atoProt = 839
        If Valor <= Aponta!Faixa13 And Valor > Aponta!Faixa12 Then txtValorCustas = FormatCurrency(Aponta!Pago7213, 2): vPago = Aponta!sPago13: atoPago = 774: atoCanc = 907: vProt = Aponta!sProt13: atoProt = 840
        If Valor <= Aponta!Faixa14 And Valor > Aponta!Faixa13 Then txtValorCustas = FormatCurrency(Aponta!Pago7214, 2): vPago = Aponta!sPago14: atoPago = 775: atoCanc = 908: vProt = Aponta!sProt14: atoProt = 841
        If Valor <= Aponta!Faixa15 And Valor > Aponta!Faixa14 Then txtValorCustas = FormatCurrency(Aponta!Pago7215, 2): vPago = Aponta!sPago15: atoPago = 776: atoCanc = 909: vProt = Aponta!sProt15: atoProt = 842
        If Valor <= Aponta!Faixa16 And Valor > Aponta!Faixa15 Then txtValorCustas = FormatCurrency(Aponta!Pago7216, 2): vPago = Aponta!sPago16: atoPago = 777: atoCanc = 910: vProt = Aponta!sProt16: atoProt = 843
        If Valor <= Aponta!Faixa17 And Valor > Aponta!Faixa16 Then txtValorCustas = FormatCurrency(Aponta!Pago7217, 2): vPago = Aponta!sPago17: atoPago = 778: atoCanc = 911: vProt = Aponta!sProt17: atoProt = 844
        If Valor <= Aponta!Faixa18 And Valor > Aponta!Faixa17 Then txtValorCustas = FormatCurrency(Aponta!Pago7218, 2): vPago = Aponta!sPago18: atoPago = 779: atoCanc = 912: vProt = Aponta!sProt18: atoProt = 845
        If Valor <= Aponta!Faixa19 And Valor > Aponta!Faixa18 Then txtValorCustas = FormatCurrency(Aponta!Pago7219, 2): vPago = Aponta!sPago19: atoPago = 780: atoCanc = 913: vProt = Aponta!sProt19: atoProt = 846
        If Valor <= Aponta!Faixa20 And Valor > Aponta!Faixa19 Then txtValorCustas = FormatCurrency(Aponta!Pago7220, 2): vPago = Aponta!sPago20: atoPago = 781: atoCanc = 914: vProt = Aponta!sProt20: atoProt = 847
        If Valor <= Aponta!Faixa21 And Valor > Aponta!Faixa20 Then txtValorCustas = FormatCurrency(Aponta!Pago7221, 2): vPago = Aponta!sPago21: atoPago = 782: atoCanc = 915: vProt = Aponta!sProt21: atoProt = 848
        If Valor <= Aponta!Faixa22 And Valor > Aponta!Faixa21 Then txtValorCustas = FormatCurrency(Aponta!Pago7222, 2): vPago = Aponta!sPago22: atoPago = 783: atoCanc = 916: vProt = Aponta!sProt22: atoProt = 849
        If Valor <= Aponta!Faixa23 And Valor > Aponta!Faixa22 Then txtValorCustas = FormatCurrency(Aponta!Pago7223, 2): vPago = Aponta!sPago23: atoPago = 784: atoCanc = 917: vProt = Aponta!sProt23: atoProt = 850
        If Valor <= Aponta!Faixa24 And Valor > Aponta!Faixa23 Then txtValorCustas = FormatCurrency(Aponta!Pago7224, 2): vPago = Aponta!sPago24: atoPago = 785: atoCanc = 918: vProt = Aponta!sProt24: atoProt = 851
        If Valor <= Aponta!Faixa25 And Valor > Aponta!Faixa24 Then txtValorCustas = FormatCurrency(Aponta!Pago7225, 2): vPago = Aponta!sPago25: atoPago = 786: atoCanc = 919: vProt = Aponta!sProt25: atoProt = 852
        If Valor <= Aponta!Faixa26 And Valor > Aponta!Faixa25 Then txtValorCustas = FormatCurrency(Aponta!Pago7226, 2): vPago = Aponta!sPago26: atoPago = 787: atoCanc = 920: vProt = Aponta!sProt26: atoProt = 853
        If Valor <= Aponta!Faixa27 And Valor > Aponta!Faixa26 Then txtValorCustas = FormatCurrency(Aponta!Pago7227, 2): vPago = Aponta!sPago27: atoPago = 788: atoCanc = 921: vProt = Aponta!sProt27: atoProt = 854
        If Valor <= Aponta!Faixa28 And Valor > Aponta!Faixa27 Then txtValorCustas = FormatCurrency(Aponta!Pago7228, 2): vPago = Aponta!sPago28: atoPago = 789: atoCanc = 922: vProt = Aponta!sProt28: atoProt = 855
        If Valor <= Aponta!Faixa29 And Valor > Aponta!Faixa28 Then txtValorCustas = FormatCurrency(Aponta!Pago7229, 2): vPago = Aponta!sPago29: atoPago = 790: atoCanc = 923: vProt = Aponta!sProt29: atoProt = 856
        If Valor <= Aponta!Faixa30 And Valor > Aponta!Faixa29 Then txtValorCustas = FormatCurrency(Aponta!Pago7230, 2): vPago = Aponta!sPago30: atoPago = 791: atoCanc = 924: vProt = Aponta!sProt30: atoProt = 857
        If Valor <= Aponta!Faixa31 And Valor > Aponta!Faixa30 Then txtValorCustas = FormatCurrency(Aponta!Pago7231, 2): vPago = Aponta!sPago31: atoPago = 792: atoCanc = 925: vProt = Aponta!sProt31: atoProt = 858
        If Valor <= Aponta!Faixa32 And Valor > Aponta!Faixa31 Then txtValorCustas = FormatCurrency(Aponta!Pago7232, 2): vPago = Aponta!sPago32: atoPago = 793: atoCanc = 926: vProt = Aponta!sProt32: atoProt = 859
        If Valor <= Aponta!Faixa33 And Valor > Aponta!Faixa32 Then txtValorCustas = FormatCurrency(Aponta!Pago7233, 2): vPago = Aponta!sPago33: atoPago = 794: atoCanc = 927: vProt = Aponta!sProt33: atoProt = 860
        If Valor <= Aponta!Faixa34 And Valor > Aponta!Faixa33 Then txtValorCustas = FormatCurrency(Aponta!Pago7234, 2): vPago = Aponta!sPago34: atoPago = 795: atoCanc = 928: vProt = Aponta!sProt34: atoProt = 861
        If Valor <= Aponta!Faixa35 And Valor > Aponta!Faixa34 Then txtValorCustas = FormatCurrency(Aponta!Pago7235, 2): vPago = Aponta!sPago35: atoPago = 796: atoCanc = 929: vProt = Aponta!sProt35: atoProt = 862
        If Valor <= Aponta!Faixa36 And Valor > Aponta!Faixa35 Then txtValorCustas = FormatCurrency(Aponta!Pago7236, 2): vPago = Aponta!sPago36: atoPago = 797: atoCanc = 930: vProt = Aponta!sProt36: atoProt = 863
        If Valor <= Aponta!Faixa37 And Valor > Aponta!Faixa36 Then txtValorCustas = FormatCurrency(Aponta!Pago7237, 2): vPago = Aponta!sPago37: atoPago = 798: atoCanc = 931: vProt = Aponta!sProt37: atoProt = 864
        If Valor <= Aponta!Faixa38 And Valor > Aponta!Faixa37 Then txtValorCustas = FormatCurrency(Aponta!Pago7238, 2): vPago = Aponta!sPago38: atoPago = 799: atoCanc = 932: vProt = Aponta!sProt38: atoProt = 865
        If Valor <= Aponta!Faixa39 And Valor > Aponta!Faixa38 Then txtValorCustas = FormatCurrency(Aponta!Pago7239, 2): vPago = Aponta!sPago39: atoPago = 800: atoCanc = 933: vProt = Aponta!sProt39: atoProt = 866
        If Valor <= Aponta!Faixa40 And Valor > Aponta!Faixa39 Then txtValorCustas = FormatCurrency(Aponta!Pago7240, 2): vPago = Aponta!sPago40: atoPago = 801: atoCanc = 934: vProt = Aponta!sProt40: atoProt = 867
        If Valor <= Aponta!Faixa41 And Valor > Aponta!Faixa40 Then txtValorCustas = FormatCurrency(Aponta!Pago7241, 2): vPago = Aponta!sPago41: atoPago = 802: atoCanc = 935: vProt = Aponta!sProt41: atoProt = 868
        If Valor <= Aponta!Faixa42 And Valor > Aponta!Faixa41 Then txtValorCustas = FormatCurrency(Aponta!Pago7242, 2): vPago = Aponta!sPago42: atoPago = 803: atoCanc = 936: vProt = Aponta!sProt42: atoProt = 869
        If Valor <= Aponta!Faixa43 And Valor > Aponta!Faixa42 Then txtValorCustas = FormatCurrency(Aponta!Pago7243, 2): vPago = Aponta!sPago43: atoPago = 804: atoCanc = 937: vProt = Aponta!sProt43: atoProt = 870
        If Valor <= Aponta!Faixa44 And Valor > Aponta!Faixa43 Then txtValorCustas = FormatCurrency(Aponta!Pago7244, 2): vPago = Aponta!sPago44: atoPago = 805: atoCanc = 938: vProt = Aponta!sProt44: atoProt = 871
        If Valor <= Aponta!Faixa45 And Valor > Aponta!Faixa44 Then txtValorCustas = FormatCurrency(Aponta!Pago7245, 2): vPago = Aponta!sPago45: atoPago = 806: atoCanc = 939: vProt = Aponta!sProt45: atoProt = 872
        If Valor <= Aponta!Faixa46 And Valor > Aponta!Faixa45 Then txtValorCustas = FormatCurrency(Aponta!Pago7246, 2): vPago = Aponta!sPago46: atoPago = 807: atoCanc = 940: vProt = Aponta!sProt46: atoProt = 873
        If Valor <= Aponta!Faixa47 And Valor > Aponta!Faixa46 Then txtValorCustas = FormatCurrency(Aponta!Pago7247, 2): vPago = Aponta!sPago47: atoPago = 808: atoCanc = 941: vProt = Aponta!sProt47: atoProt = 874
        If Valor <= Aponta!Faixa48 And Valor > Aponta!Faixa47 Then txtValorCustas = FormatCurrency(Aponta!Pago7248, 2): vPago = Aponta!sPago48: atoPago = 809: atoCanc = 942: vProt = Aponta!sProt48: atoProt = 875
        If Valor <= Aponta!Faixa49 And Valor > Aponta!Faixa48 Then txtValorCustas = FormatCurrency(Aponta!Pago7249, 2): vPago = Aponta!sPago49: atoPago = 810: atoCanc = 943: vProt = Aponta!sProt49: atoProt = 876
        If Valor <= Aponta!Faixa50 And Valor > Aponta!Faixa49 Then txtValorCustas = FormatCurrency(Aponta!Pago7250, 2): vPago = Aponta!sPago50: atoPago = 811: atoCanc = 944: vProt = Aponta!sProt50: atoProt = 877
        If Valor <= Aponta!Faixa51 And Valor > Aponta!Faixa50 Then txtValorCustas = FormatCurrency(Aponta!Pago7251, 2): vPago = Aponta!sPago51: atoPago = 812: atoCanc = 945: vProt = Aponta!sProt51: atoProt = 878
        If Valor <= Aponta!Faixa52 And Valor > Aponta!Faixa51 Then txtValorCustas = FormatCurrency(Aponta!Pago7252, 2): vPago = Aponta!sPago52: atoPago = 813: atoCanc = 946: vProt = Aponta!sProt52: atoProt = 879
        If Valor <= Aponta!Faixa53 And Valor > Aponta!Faixa52 Then txtValorCustas = FormatCurrency(Aponta!Pago7253, 2): vPago = Aponta!sPago53: atoPago = 814: atoCanc = 947: vProt = Aponta!sProt53: atoProt = 880
        If Valor <= Aponta!Faixa54 And Valor > Aponta!Faixa53 Then txtValorCustas = FormatCurrency(Aponta!Pago7254, 2): vPago = Aponta!sPago54: atoPago = 815: atoCanc = 948: vProt = Aponta!sProt54: atoProt = 881
        If Valor <= Aponta!Faixa55 And Valor > Aponta!Faixa54 Then txtValorCustas = FormatCurrency(Aponta!Pago7255, 2): vPago = Aponta!sPago55: atoPago = 816: atoCanc = 949: vProt = Aponta!sProt55: atoProt = 882
        If Valor <= Aponta!Faixa56 And Valor > Aponta!Faixa55 Then txtValorCustas = FormatCurrency(Aponta!Pago7256, 2): vPago = Aponta!sPago56: atoPago = 817: atoCanc = 950: vProt = Aponta!sProt56: atoProt = 883
        If Valor <= Aponta!Faixa57 And Valor > Aponta!Faixa56 Then txtValorCustas = FormatCurrency(Aponta!Pago7257, 2): vPago = Aponta!sPago57: atoPago = 818: atoCanc = 951: vProt = Aponta!sProt57: atoProt = 884
        If Valor <= Aponta!Faixa58 And Valor > Aponta!Faixa57 Then txtValorCustas = FormatCurrency(Aponta!Pago7258, 2): vPago = Aponta!sPago58: atoPago = 819: atoCanc = 952: vProt = Aponta!sProt58: atoProt = 885
        If Valor <= Aponta!Faixa59 And Valor > Aponta!Faixa58 Then txtValorCustas = FormatCurrency(Aponta!Pago7259, 2): vPago = Aponta!sPago59: atoPago = 820: atoCanc = 953: vProt = Aponta!sProt59: atoProt = 886
        If Valor <= Aponta!Faixa60 And Valor > Aponta!Faixa59 Then txtValorCustas = FormatCurrency(Aponta!Pago7260, 2): vPago = Aponta!sPago60: atoPago = 821: atoCanc = 954: vProt = Aponta!sProt60: atoProt = 887
        If Valor <= Aponta!Faixa61 And Valor > Aponta!Faixa60 Then txtValorCustas = FormatCurrency(Aponta!Pago7261, 2): vPago = Aponta!sPago61: atoPago = 822: atoCanc = 955: vProt = Aponta!sProt61: atoProt = 888
        If Valor <= Aponta!Faixa62 And Valor > Aponta!Faixa61 Then txtValorCustas = FormatCurrency(Aponta!Pago7262, 2): vPago = Aponta!sPago62: atoPago = 823: atoCanc = 956: vProt = Aponta!sProt62: atoProt = 889
        If Valor <= Aponta!Faixa63 And Valor > Aponta!Faixa62 Then txtValorCustas = FormatCurrency(Aponta!Pago7263, 2): vPago = Aponta!sPago63: atoPago = 824: atoCanc = 957: vProt = Aponta!sProt63: atoProt = 890
        If Valor <= Aponta!Faixa64 And Valor > Aponta!Faixa63 Then txtValorCustas = FormatCurrency(Aponta!Pago7264, 2): vPago = Aponta!sPago64: atoPago = 825: atoCanc = 958: vProt = Aponta!sProt64: atoProt = 891
        If Valor > Aponta!Faixa64 Then txtValorCustas = FormatCurrency(Aponta!Pago7265, 2): vPago = Aponta!sPago65: atoPago = 826: atoCanc = 959: vProt = Aponta!sProt65: atoProt = 892
End Sub


