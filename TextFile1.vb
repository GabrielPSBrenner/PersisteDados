''''Class LerDiretoSQL

''''Public Function TabelaSQLServer(ByRef Conexao As SqlClient.SqlConnection, ByVal Tabela As String) As stTabela
''''   Return MontaTabela(Conexao, Tabela)
''''End Function

''''Public Function TabelaSQLServer(ByVal SQLServer As String, ByVal Database As String, ByVal UID As String, ByVal PWD As String, ByVal Tabela As String) As stTabela
''''   Dim Cn As New SqlClient.SqlConnection("Server=" & SQLServer & ";Database=" & Database & ";UID=" & UID & ";PWD=" & PWD)
''''   Dim lstabela As stTabela = MontaTabela(Cn, Tabela)
''''   Cn.Close()
''''   Cn.Dispose()
''''   Cn = Nothing
''''   Return lstabela
''''End Function

''''Private Function MontaTabela(ByRef Cn As SqlClient.SqlConnection, ByVal Tabela As String) As stTabela
''''   Dim SQLPk As String = "select  distinct isnull(syscolumns.name, '') as ColumnName from  sysobjects,  sysindexes,  sysindexkeys, syscolumns where  sysobjects.name = '" & Tabela & "' and sysindexes.id = sysobjects.id and sysindexes.status & 2 = 2 and sysindexkeys.id = sysindexes.id and sysindexkeys.indid = sysindexes.indid and syscolumns.id = sysindexkeys.id and syscolumns.colid = sysindexkeys.colid"
''''   Dim oCmm As New SqlClient.SqlCommand("Sp_Columns", Cn)
''''   oCmm.CommandType = CommandType.StoredProcedure
''''   oCmm.Parameters.Clear()
''''   oCmm.Parameters.Add(New SqlClient.SqlParameter("@TAble_nAme", SqlDbType.VarChar))
''''   oCmm.Parameters("@TAble_Name").Value = Tabela
''''   Dim Da As New SqlClient.SqlDataAdapter(oCmm)
''''   'Dim Da As New SqlClient.SqlDataAdapter("sp_columns " & Tabela, Cn)
''''   Dim Dt As New DataTable()
''''   Dim DtPk As New DataTable()
''''   Da.Fill(Dt)
''''   lTabela.TotalCampos = Dt.Rows.Count
''''   '////////////////////////////
''''   'Busca os nomes dos campos pertencentes a chave primária da tabela passada
''''   '
''''   Da = New SqlClient.SqlDataAdapter(SQLPk, Cn)
''''   Da.Fill(DtPk)
''''   lTabela.TotalCamposChaves = Dt.Rows.Count

''''   lTabela.Nome = Tabela
''''   If Dt.Rows.Count > 0 Then
''''      ReDim lTabela.Campos(Dt.Rows.Count - 1)
''''   End If
''''   If DtPk.Rows.Count > 0 Then
''''      ReDim lTabela.ChavePrimaria(DtPk.Rows.Count - 1)
''''   End If
''''   For i = 0 To DtPk.Rows.Count - 1
''''      lTabela.ChavePrimaria(i).Nome = DtPk.Rows(i).Item("ColumnName").ToString
''''   Next
''''   '
''''   '****************************
''''   '////////////////////////////
''''   'Busca todos os campos da tabela passada e suas propriedades
''''   'Faz relação com as chaves primárias e atribui os valores faltando as mesmas
''''   For i = 0 To Dt.Rows.Count - 1
''''      lTabela.Campos(i).Nome = Dt.Rows(i).Item("COLUMN_NAME").ToString
''''      lTabela.Campos(i).Index = Dt.Rows(i).Item("ORDINAL_POSITION").ToString
''''      lTabela.Campos(i).Obrigatorio = (Dt.Rows(i).Item("NULLABLE").ToString = "0")
''''      lTabela.Campos(i).Tipo = FieldTypeSQLServer(Dt.Rows(i).Item("Data_Type").ToString)
''''      For ii = 0 To DtPk.Rows.Count - 1
''''         If lTabela.Campos(i).Nome = lTabela.ChavePrimaria(ii).Nome Then
''''            lTabela.Campos(i).Chave = True
''''            lTabela.ChavePrimaria(ii).Index = lTabela.Campos(i).Index
''''            lTabela.ChavePrimaria(ii).Obrigatorio = True
''''            lTabela.ChavePrimaria(ii).Tipo = lTabela.Campos(i).Tipo
''''         End If
''''      Next
''''   Next
''''   '
''''   '****************************
''''   '////////////////////////////
''''   'Libera os objetos
''''   Dt.Dispose()
''''   Da.Dispose()
''''   Da = Nothing
''''   Return lTabela
''''   '****************************
''''End Function
''''Private Function FieldTypeSQLServer(ByVal Tipo As Integer) As EnumTipo
''''   '---------------------------------------------------
''''   'valores correspondentes aos campos do Sql-Server
''''   Dim SQLSNumerico() As Integer = {-6, -11, 2, 3, 4, 5, 7}
''''   Dim SQLSTexto() As Integer = {-10, -9, -8, -1, 1, 12}
''''   Dim SQLSDataHora() As Integer = {-2, 11}
''''   Dim SQLSBinario() As Integer = {-4, -3, -2}
''''   Dim SQLSBooleano As Integer = -7
''''   '---------------------------------------------------
''''   Dim FindIndex As Integer
''''   FindIndex = Array.BinarySearch(SQLSNumerico, Tipo)
''''   If FindIndex > -1 Then
''''      Return EnumTipo.Numerico
''''   End If
''''   FindIndex = Array.BinarySearch(SQLSTexto, Tipo)
''''   If FindIndex > -1 Then
''''      Return EnumTipo.Texto
''''   End If
''''   FindIndex = Array.BinarySearch(SQLSDataHora, Tipo)
''''   If FindIndex > -1 Then
''''      Return EnumTipo.Data
''''   End If
''''   FindIndex = Array.BinarySearch(SQLSBinario, Tipo)
''''   If FindIndex > -1 Then
''''      Return EnumTipo.Binario
''''   End If
''''   If Tipo = SQLSBooleano Then
''''      Return EnumTipo.Booleano
''''   End If
''''End Function

''''Private Function FieldTypeSQLServer(ByVal Tipo As Integer) As EnumTipo
''''   '---------------------------------------------------
''''   'valores correspondentes aos campos do Sql-Server
''''   Dim SQLSNumerico() As Integer = {-6, -11, 2, 3, 4, 5, 7}
''''   Dim SQLSTexto() As Integer = {-10, -9, -8, -1, 1, 12}
''''   Dim SQLSDataHora() As Integer = {-2, 11}
''''   Dim SQLSBinario() As Integer = {-4, -3, -2}
''''   Dim SQLSBooleano As Integer = -7
''''   '---------------------------------------------------
''''   Dim FindIndex As Integer
''''   FindIndex = Array.BinarySearch(SQLSNumerico, Tipo)
''''   If FindIndex > -1 Then
''''      Return EnumTipo.Numerico
''''   End If
''''   FindIndex = Array.BinarySearch(SQLSTexto, Tipo)
''''   If FindIndex > -1 Then
''''      Return EnumTipo.Texto
''''   End If
''''   FindIndex = Array.BinarySearch(SQLSDataHora, Tipo)
''''   If FindIndex > -1 Then
''''      Return EnumTipo.Data
''''   End If
''''   FindIndex = Array.BinarySearch(SQLSBinario, Tipo)
''''   If FindIndex > -1 Then
''''      Return EnumTipo.Binario
''''   End If
''''   If Tipo = SQLSBooleano Then
''''      Return EnumTipo.Booleano
''''   End If
''''End Function
''''Private Function MontaTabela(ByRef Cn As SqlClient.SqlConnection, ByVal Tabela As String) As stTabela
''''   Dim SQLPk As String = "select  distinct isnull(syscolumns.name, '') as ColumnName from  sysobjects,  sysindexes,  sysindexkeys, syscolumns where  sysobjects.name = '" & Tabela & "' and sysindexes.id = sysobjects.id and sysindexes.status & 2 = 2 and sysindexkeys.id = sysindexes.id and sysindexkeys.indid = sysindexes.indid and syscolumns.id = sysindexkeys.id and syscolumns.colid = sysindexkeys.colid"
''''   Dim oCmm As New SqlClient.SqlCommand("Sp_Columns", Cn)
''''   oCmm.CommandType = CommandType.StoredProcedure
''''   oCmm.Parameters.Clear()
''''   oCmm.Parameters.Add(New SqlClient.SqlParameter("@TAble_nAme", SqlDbType.VarChar))
''''   oCmm.Parameters("@TAble_Name").Value = Tabela
''''   Dim Da As New SqlClient.SqlDataAdapter(oCmm)
''''   'Dim Da As New SqlClient.SqlDataAdapter("sp_columns " & Tabela, Cn)
''''   Dim Dt As New DataTable()
''''   Dim DtPk As New DataTable()
''''   Da.Fill(Dt)
''''   lTabela.TotalCampos = Dt.Rows.Count
''''   '////////////////////////////
''''   'Busca os nomes dos campos pertencentes a chave primária da tabela passada
''''   '
''''   Da = New SqlClient.SqlDataAdapter(SQLPk, Cn)
''''   Da.Fill(DtPk)
''''   lTabela.TotalCamposChaves = Dt.Rows.Count

''''   lTabela.Nome = Tabela
''''   If Dt.Rows.Count > 0 Then
''''      ReDim lTabela.Campos(Dt.Rows.Count - 1)
''''   End If
''''   If DtPk.Rows.Count > 0 Then
''''      ReDim lTabela.ChavePrimaria(DtPk.Rows.Count - 1)
''''   End If
''''   For i = 0 To DtPk.Rows.Count - 1
''''      lTabela.ChavePrimaria(i).Nome = DtPk.Rows(i).Item("ColumnName").ToString
''''   Next
''''   '
''''   '****************************
''''   '////////////////////////////
''''   'Busca todos os campos da tabela passada e suas propriedades
''''   'Faz relação com as chaves primárias e atribui os valores faltando as mesmas
''''   For i = 0 To Dt.Rows.Count - 1
''''      lTabela.Campos(i).Nome = Dt.Rows(i).Item("COLUMN_NAME").ToString
''''      lTabela.Campos(i).Index = Dt.Rows(i).Item("ORDINAL_POSITION").ToString
''''      lTabela.Campos(i).Obrigatorio = (Dt.Rows(i).Item("NULLABLE").ToString = "0")
''''      lTabela.Campos(i).Tipo = FieldTypeSQLServer(Dt.Rows(i).Item("Data_Type").ToString)
''''      For ii = 0 To DtPk.Rows.Count - 1
''''         If lTabela.Campos(i).Nome = lTabela.ChavePrimaria(ii).Nome Then
''''            lTabela.Campos(i).Chave = True
''''            lTabela.ChavePrimaria(ii).Index = lTabela.Campos(i).Index
''''            lTabela.ChavePrimaria(ii).Obrigatorio = True
''''            lTabela.ChavePrimaria(ii).Tipo = lTabela.Campos(i).Tipo
''''         End If
''''      Next
''''   Next
''''   '
''''   '****************************
''''   '////////////////////////////
''''   'Libera os objetos
''''   Dt.Dispose()
''''   Da.Dispose()
''''   Da = Nothing
''''      Return lTabela
''''   End Function      '****************************

''''End Class
