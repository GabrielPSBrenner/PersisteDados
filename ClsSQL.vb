Imports System.Reflection
Imports System.EnterpriseServices

Public Class ClsSQL

   Public Enum EnumType 'Enumerador que identifica o tipo de campo 
      eFloat = 0
      eInteger = 1
      eText = 2
      eDate = 3
      eTime = 4
      eBin = 5
      eBool = 6
      eDateTime = 7 'Falta identificar o tipo data/hora junto <<IMPLEMENTAR>>
   End Enum

   Public Enum EnumDBType 'Enumerador que identifica o tipo de servidor de dados que será acessado
      SQLServer = 0
      Oracle = 1
      MySQL = 2
      OLEDB = 3
      ODBC = 4
   End Enum

   Public Structure stField 'Estrutura criada, que representa os campos dos objetos
      Dim Key As Boolean
      Dim FK As Boolean
      Dim AutoNumera As Boolean
      Dim ObjFieldName As String
      Dim TableName As String
      Dim TableFieldName As String
      Dim Type As EnumType
      Dim Value As Object
   End Structure

   Private lDataBaseType As EnumDBType 'Variável interna que identifica o tipo de banco de dados que está sendo acessado

   Public Function SQLInsert(ByVal TableName As String, ByVal Prefix As String, ByVal Field() As stField) As String
      '************************************************************
      'Monta a expressão SQL de inserção a partir da tabela passada
      'Retorna a expressão SQL que será utilizada para inserir o registro no banco de dados
      '-------------PARAMETROS RECEBIDOS-------------
      'TableName -> Nome da tabela de referência para montar o SQL
      'Field()   -> Coleção de stField, representando os campos na tabela e as propriedades nos objetos
      '---------FIM DOS PARAMETROS RECEBIDOS-------------            
      '************************************************************
      Dim i As Integer
      Dim sInsert As New System.Text.StringBuilder
      Dim sValues As New System.Text.StringBuilder
      sInsert.Append("Insert into ")
      sInsert.Append(Prefix)
      sInsert.Append(TableName)
      sInsert.Append("(")
      sValues.Append("(")
      For i = 0 To Field.Length - 1
         If Field(i).AutoNumera = False Then
            sInsert.Append(Field(i).TableFieldName)
            If Field(i).Type = EnumType.eDate Then
               If IsDate(Field(i).Value) Then
                  If CDate(Field(i).Value) = CDate("0001-01-01 00:00:00.000") Then
                     sValues.Append("NULL")
                  Else
                     If lDataBaseType <> EnumDBType.OLEDB Then
                        sValues.Append("'" & Format(CDate(Field(i).Value), "yyyy/MM/dd") & "'")
                     Else
                        sValues.Append("'" & Format(CDate(Field(i).Value), "MM/dd/yyyy") & "'")
                     End If
                  End If
               Else
                  sValues.Append("NULL")
               End If
            ElseIf Field(i).Type = EnumType.eTime Then
               If IsDate(Field(i).Value) Then
                  sValues.Append("'" & Format(CDate(Field(i).Value), "HH:mm:ss") & "'")
               Else
                  sValues.Append("NULL")
               End If
            ElseIf Field(i).Type = EnumType.eDateTime Then
               If IsDate(Field(i).Value) Then
                  If CDate(Field(i).Value) = CDate("0001-01-01 00:00:00.000") Then
                     sValues.Append("NULL")
                  Else
                     If lDataBaseType <> EnumDBType.OLEDB Then
                        sValues.Append("'" & Format(CDate(Field(i).Value), "yyyy-MM-dd") & " " & Format(CDate(Field(i).Value), "HH:mm:ss") & ".000'")
                     Else
                        sValues.Append("#" & Format(CDate(Field(i).Value), "MM/dd/yyyy") & " " & Format(CDate(Field(i).Value), "HH:mm:ss") & "#")
                     End If
                  End If
               Else
                  sValues.Append("NULL")
               End If
            ElseIf Field(i).Type = EnumType.eBool Then
               If Field(i).Value.ToString = "-1" Or Field(i).Value.ToString = "1" Or Field(i).Value = True Then
                  sValues.Append("1")
               Else
                  sValues.Append("0")
               End If
            ElseIf Field(i).Type = EnumType.eInteger Then
               If Field(i).FK And Field(i).Value = 0 Then
                  sValues.Append("NULL")
               Else
                  sValues.Append(Field(i).Value.ToString)
               End If
            ElseIf Field(i).Type = EnumType.eFloat Then
               If Field(i).FK And Field(i).Value = 0 Then
                  sValues.Append("NULL")
               Else
                  sValues.Append(Field(i).Value.ToString.Replace(",", "."))
               End If
            ElseIf Field(i).Type = EnumType.eText Then
               'The replace is to prevent SQLInjection statement
               If IsNothing(Field(i).Value) Then
                  sValues.Append("NULL")
               Else
                  If Field(i).Value.ToString.Trim = "" Then
                     sValues.Append("NULL")
                  Else
                     sValues.Append("'" & Field(i).Value.ToString.Replace("'", "") & "'")
                  End If
               End If
            End If
            If i < Field.Length - 1 Then
               sInsert.Append(",")
               sValues.Append(",")
            End If
         End If
      Next
      sInsert.Append(") values")
      sValues.Append(")")
      sInsert.Append(sValues.ToString)
      Return sInsert.ToString
   End Function

   Public Function SQLUpdate(ByVal TableName As String, ByVal Prefix As String, ByVal Field() As stField) As String
      '************************************************************
      'Monta a expressão SQL de atualização a partir da tabela passada
      'Retorna a expressão SQL que será utilizada para atualizar o registro no banco de dados
      'É obrigatório, para atualização, a tabela possuir campos identificadores únicos,
      'representando a chave primária e, tal informação estar configurada no XML
      '-------------PARAMETROS RECEBIDOS-------------
      'TableName -> Nome da tabela de referência para montar o SQL
      'Field()   -> Coleção de stField, representando os campos na tabela e as propriedades nos objetos
      '---------FIM DOS PARAMETROS RECEBIDOS-------------            
      '************************************************************

      Dim i As Integer
      Dim sUpdate As New System.Text.StringBuilder("Update ")
      sUpdate.Append(Prefix)
      sUpdate.Append(TableName)
      sUpdate.Append(" set ")
      For i = 0 To Field.Length - 1
         If Not (Field(i).AutoNumera) Then
            'If Not (Field(i).AutoNumera Or Field(i).Key) Then
            sUpdate.Append(Field(i).TableFieldName)
            sUpdate.Append("=")
            '*****************************************************************
            'Inicia a construção da segunda parte do SQL, que contém os valores dos campos
            'Preocupa-se com os tipos bases, Data, Booleano, String e Numérico
            If Field(i).Type = EnumType.eDate Then
               If Field(i).Value.ToString.Trim <> "" Then
                  If IsDate(Field(i).Value) Then
                     If CDate(Field(i).Value) = CDate("0001-01-01 00:00:00.000") Then
                        sUpdate.Append("NULL")
                     Else
                        If lDataBaseType <> EnumDBType.OLEDB Then
                           sUpdate.Append("'" & Format(CDate(Field(i).Value), "yyyy/MM/dd") & "'")
                        Else
                           sUpdate.Append("'" & Format(CDate(Field(i).Value), "MM/dd/yyyy") & "'")
                        End If
                     End If
                  Else
                     sUpdate.Append("NULL")
                  End If
               Else
                  sUpdate.Append("NULL")
               End If
            ElseIf Field(i).Type = EnumType.eTime Then
               If IsDate(Field(i).Value) Then
                  sUpdate.Append("'" & Format(CDate(Field(i).Value), "HH:mm:ss") & "'")
               Else
                  sUpdate.Append("NULL")
               End If
            ElseIf Field(i).Type = EnumType.eDateTime Then
               If IsDate(Field(i).Value) Then
                  If CDate(Field(i).Value) = CDate("0001-01-01 00:00:00.000") Then
                     sUpdate.Append("NULL")
                  Else
                     If lDataBaseType <> EnumDBType.OLEDB Then
                        sUpdate.Append("'" & Format(CDate(Field(i).Value), "yyyy-MM-dd") & " " & Format(CDate(Field(i).Value), "HH:mm:ss") & ".000'")
                     Else
                        sUpdate.Append("#" & Format(CDate(Field(i).Value), "MM/dd/yyyy") & " " & Format(CDate(Field(i).Value), "HH:mm:ss") & "#")
                     End If
                  End If
               Else
                  sUpdate.Append("NULL")
               End If
            ElseIf Field(i).Type = EnumType.eBool Then
               If Field(i).Value.ToString = "-1" Or Field(i).Value.ToString.ToLower = "true" Or Field(i).Value = 1 Then
                  sUpdate.Append("1")
               Else
                  sUpdate.Append("0")
               End If
            ElseIf Field(i).Type = EnumType.eInteger Then
               If Field(i).FK And Field(i).Value = 0 Then
                  sUpdate.Append("NULL")
               Else
                  sUpdate.Append(Field(i).Value.ToString)
               End If
            ElseIf Field(i).Type = EnumType.eFloat Then
               If Field(i).FK And Field(i).Value = 0 Then
                  sUpdate.Append("NULL")
               Else
                  sUpdate.Append(Field(i).Value.ToString.Replace(",", "."))
               End If
            ElseIf Field(i).Type = EnumType.eText Then
               If IsNothing(Field(i).Value) Then
                  sUpdate.Append("NULL")
               Else
                  If Field(i).Value.ToString.Trim = "" Then
                     sUpdate.Append("NULL")
                  Else
                     'The replace is to prevent SQLInjection statement
                     sUpdate.Append("'" & Replace(Field(i).Value, "'", "") & "'")
                  End If
               End If
            End If
            If i < Field.Length - 1 Then
               sUpdate.Append(",")
            End If
         End If
      Next
      sUpdate.Append(" ")
      sUpdate.Append(SQLWhere(Field))
      Return sUpdate.ToString
   End Function

   Public Function SQLDelete(ByVal TableName As String, ByVal Prefix As String, ByVal Field() As stField) As String
      '************************************************************
      'Monta a expressão SQL de exclusão a partir da tabela passada
      'Retorna a expressão SQL que será utilizada para excluir o registro no banco de dados
      'É obrigatório para exclusão, a tabela possuir campos identificadores únicos,
      'representando a chave primária e, tal informação estar configurada no XML
      '-------------PARAMETROS RECEBIDOS-------------
      'TableName -> Nome da tabela de referência para montar o SQL
      'Field()   -> Coleção de stField, representando os campos na tabela e as propriedades nos objetos
      '---------FIM DOS PARAMETROS RECEBIDOS-------------      
      '************************************************************

      Dim sDelete As New System.Text.StringBuilder("Delete ")
      If lDataBaseType = EnumDBType.OLEDB Then
         sDelete.Append("* from ")
      Else
         sDelete.Append("from ")
      End If
      sDelete.Append(Prefix)
      sDelete.Append(TableName)
      sDelete.Append(" ")
      sDelete.Append(SQLWhere(Field))
      Return sDelete.ToString
   End Function

   Public Function SQLDeleteCriteria(ByVal TableName As String, ByVal Prefix As String, ByVal Criteria As ClsCriteria) As String
      '************************************************************
      'Monta a expressão SQL de exclusão a partir da tabela passada
      'Retorna a expressão SQL que será utilizada para excluir os registros no banco de dados
      '-------------PARAMETROS RECEBIDOS-------------
      'TableName -> Nome da tabela de referência para montar o SQL
      'Criteria  -> Critério para a exclusão dos registros
      '---------FIM DOS PARAMETROS RECEBIDOS-------------      
      '************************************************************

      Dim sDelete As New System.Text.StringBuilder("Delete ")
      If lDataBaseType = EnumDBType.OLEDB Then
         sDelete.Append("* from ")
      Else
         sDelete.Append("from ")
      End If
      sDelete.Append(Prefix)
      sDelete.Append(TableName)
      sDelete.Append(" ")
      sDelete.Append(Criteria.SQLWhere)
      Return sDelete.ToString
   End Function

   Public Function SQLWhere(ByVal Field() As stField, Optional ByVal LikeUse As Boolean = False, Optional ByVal OnlyPK As Boolean = True) As String
      '*****************************************************
      'Monta uma expressão de Where, a partir de uma coleção de stField passada
      'Serve para seleções simples, linkadas diretamente ao objeto, com poucos critérios de seleção
      'Será utilizada para que a classe monte os critérios de atualização e exclusão
      'Elimina valores numéricos que sejam zero (porque a variável numérica se inicializa com zero, dificultando a seleção)
      '-------------PARAMETROS RECEBIDOS-------------
      'Field() -> Coleção de campos que serão utilizados para seleçao
      'LikeUse -> Determina se os campos do tipo Text utilizam like como operador ao invés do igual
      'OnlyPk  -> Determina se somente as chaves primárias participarão dos critérios de seleção 
      '---------FIM DOS PARAMETROS RECEBIDOS-------------
      '*****************************************************

      Dim sRetorno As New System.Text.StringBuilder
      Dim Enter As Boolean
      Dim i As Byte
      For i = 0 To Field.Length - 1
         If (OnlyPK = True And Field(i).Key = True) Or OnlyPK = False Then
            If Not IsNothing(Field(i).Value) Then
               'Se a propriedade for igual a nothing, não participa do critério de seleção
               'Só participa do critério strings diferente de vazio ou valores numéricos diferentes de zero
               Select Case Field(i).Type
                  Case EnumType.eBool
                     Enter = False
                  Case EnumType.eText
                     If Field(i).Value.ToString.Trim = "" Then
                        Enter = False
                     Else
                        Enter = True
                     End If
                  Case EnumType.eFloat, EnumType.eInteger
                     If Field(i).Value = 0 Then
                        Enter = False
                     Else
                        Enter = True
                     End If
                  Case EnumType.eDate, EnumType.eDateTime
                     If Field(i).Value = CDate("01/01/0001") Then
                        Enter = False
                     Else
                        Enter = True
                     End If
                  Case EnumType.eTime
                     If Field(i).Value = CDate("00:00") Then
                        Enter = False
                     Else
                        Enter = True
                     End If
                  Case Else
                     Enter = True
               End Select

               If Enter Then
                  If sRetorno.ToString.Trim <> "" Then
                     'existe um Pré-SQL
                     sRetorno.Append(" and ")
                  End If
                  sRetorno.Append(Field(i).TableFieldName)

                  If Field(i).Type = EnumType.eBool Then
                     If Field(i).Value = "-1" Or Field(i).Value = "1" Or Field(i).Value.ToString.ToLower = "true" Then
                        sRetorno.Append("=1 ")
                     Else
                        sRetorno.Append("=0 ")
                     End If
                  ElseIf Field(i).Type = EnumType.eDate Then
                     If IsDate(Field(i).Value) Then
                        If lDataBaseType <> EnumDBType.OLEDB Then
                           sRetorno.Append("='" & Format(CDate(Field(i).Value), "yyyy/MM/dd") & "'")
                        Else
                           sRetorno.Append("='" & Format(CDate(Field(i).Value), "MM/dd/yyyy") & "'")
                        End If
                     Else
                        sRetorno.Append("NULL")
                     End If
                  ElseIf Field(i).Type = EnumType.eDateTime Then
                     If IsDate(Field(i).Value) Then
                        If lDataBaseType <> EnumDBType.OLEDB Then
                           sRetorno.Append("='" & Format(CDate(Field(i).Value), "yyyy/MM/dd HH:mm:ss") & "'")
                        Else
                           sRetorno.Append("='" & Format(CDate(Field(i).Value), "MM/dd/yyyy HH:mm:ss") & "'")
                        End If
                     Else
                        sRetorno.Append("NULL")
                     End If
                  ElseIf Field(i).Type = EnumType.eTime Then
                     If IsDate(Field(i).Value) Then
                        sRetorno.Append("='" & Format(CDate(Field(i).Value), "HH:mm:ss") & "'")
                     Else
                        sRetorno.Append("NULL")
                     End If
                  ElseIf Field(i).Type = EnumType.eInteger Then
                     If IsNumeric(Field(i).Value) Then
                        sRetorno.Append("=" & Field(i).Value)
                     Else
                        sRetorno.Append("=NULL")
                     End If
                  ElseIf Field(i).Type = EnumType.eFloat Then
                     If IsNumeric(Field(i).Value) Then
                        sRetorno.Append("=" & Field(i).Value.ToString.Replace(",", "."))
                     Else
                        sRetorno.Append("=NULL")
                     End If
                  ElseIf Field(i).Type = EnumType.eText Then
                     'The replace is to prevent SQLInjection statement
                     If LikeUse Then
                        sRetorno.Append(" like ")
                        sRetorno.Append("'%" + Field(i).Value.ToString.Replace("'", "") + "%'")
                     Else
                        sRetorno.Append(" ='" + Field(i).Value.ToString.Replace("'", "") + "'")
                     End If
                  End If
               End If
            End If
         End If
      Next
      If sRetorno.ToString.Trim <> "" Then sRetorno.Insert(0, " Where ")
      Return sRetorno.ToString
   End Function

   Public Function MontSQLTop(ByVal TableName As String, ByVal TableField As String, Optional ByVal Criteria As ClsCriteria = Nothing) As String
      Dim SQL As New Text.StringBuilder
      Select Case lDataBaseType
         Case EnumDBType.SQLServer, EnumDBType.Oracle, EnumDBType.OLEDB
            SQL.Append("Select Top 1 ")
            SQL.Append(TableField)
            SQL.Append(" from ")
            SQL.Append(TableName)
            If Not IsNothing(Criteria) Then
               SQL.Append(Criteria.SQLWhere)
            End If
            SQL.Append(" order by ")
            SQL.Append(TableField)
            SQL.Append(" Desc ")
         Case Else
            SQL.Append("Select top 1 ")
            SQL.Append(TableField)
            SQL.Append(" from ")
            SQL.Append(TableName)
            If Not IsNothing(Criteria) Then
               SQL.Append(Criteria.SQLWhere)
            End If
            SQL.Append(" order by ")
            SQL.Append(TableField)
            SQL.Append(" Desc ")
      End Select
      Return SQL.ToString
   End Function

   Public Function MontSQLSingleCount(ByVal TableName As String, ByVal TableField As String, Optional ByVal Criteria As ClsCriteria = Nothing) As String
      Dim SQL As New Text.StringBuilder
      SQL.Append("Select Count(")
      SQL.Append(TableField)
      SQL.Append(") from ")
      SQL.Append(TableName)
      If Not IsNothing(Criteria) Then
         SQL.Append(Criteria.SQLWhere)
         SQL.Append(" ")
      End If
      Return SQL.ToString
   End Function

   Public Sub New(ByVal DataBaseType As EnumDBType)
      'Recebe qual tipo de banco de dados está sendo acessado
      lDataBaseType = DataBaseType
   End Sub

End Class
