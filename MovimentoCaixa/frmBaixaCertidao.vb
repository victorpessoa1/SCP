
Private Sub DTData_Change()
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim RSFinan As ADODB.Recordset
    Set RSFinan = New ADODB.Recordset
    MeuGrid
    RS.Open "SELECT * FROM tblReqCertidao WHERE Data_Req='" & Format(Me.DTData, "mm/dd/yyyy") & "'", DB, adOpenDynamic
    If RS.RecordCount > 0 Then
        Do While Not RS.EOF
            If RS!Pago = True Then
                RSFinan.Open "SELECT Cartao FROM tblfinanceiro WHERE Codigo = '" & RS!Codigo & "'", DB, adOpenDynamic
                If RSFinan!Cartao = True Then
                    Situacao = "PAGO  |  C"
                Else
                    Situacao = "PAGO  |  D"
                End If
                RSFinan.Close
            End If
            If RS!Pago = False Then Situacao = "PENDENTE"
            If RS!Impresso = True Then Impresso = "SIM"
            If RS!Impresso = False Then Impresso = "NÃO"
            If RS!CENPROT = True Then nCenprot = "SIM"
            If RS!CENPROT = False Then nCenprot = "NÃO"
            Me.Grid.AddItem RS("Codigo") & vbTab & Trim(RS("Nome")) & vbTab & Format(RS("Data_Req"), "dd/mm/yyyy") & vbTab & RS("Pagar") & vbTab & Situacao & vbTab & Impresso & vbTab & RS("Usuario") & vbTab & nCenprot
            RS.MoveNext
        Loop
    End If
    ZEBRAR
End Sub

Private Sub Form_Load()
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    Dim RSFinan As ADODB.Recordset
    Set RSFinan = New ADODB.Recordset
    Me.optDataLiquida = False
    Data_Liquida = False
    Me.DTData.Value = Date
    nCodigo = ""
    ICod = ""
    MeuGrid
    RS.Open "SELECT * FROM tblReqCertidao WHERE Data_Req='" & Format(Me.DTData, "mm/dd/yyyy") & "'ORDER BY idCertidao DESC", DB, adOpenDynamic
    If RS.RecordCount > 0 Then
        Do While Not RS.EOF
            If RS!Pago = True Then
                RSFinan.Open "SELECT Cartao FROM tblfinanceiro WHERE Codigo = '" & RS!Codigo & "'", DB, adOpenDynamic
                If RSFinan!Cartao = True Then
                    Situacao = "PAGO  |  C"
                Else
                    Situacao = "PAGO  |  D"
                End If
                RSFinan.Close
            End If
            If RS!Pago = False Then Situacao = "PENDENTE"
            If RS!Impresso = True Then Impresso = "SIM"
            If RS!Impresso = False Then Impresso = "NÃO"
            If RS!CENPROT = True Then nCenprot = "SIM"
            If RS!CENPROT = False Then nCenprot = "NÃO"
            If RS!Gratuito = True Then Situacao = "GRATUITO"
            Me.Grid.AddItem RS("Codigo") & vbTab & Trim(RS("Nome")) & vbTab & Format(RS("Data_Req"), "dd/mm/yyyy") & vbTab & RS("Pagar") & vbTab & Situacao & vbTab & Impresso & vbTab & RS("Usuario") & vbTab & nCenprot
            RS.MoveNext
        Loop
    End If
    ZEBRAR
End Sub