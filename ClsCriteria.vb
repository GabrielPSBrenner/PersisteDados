Imports System.Reflection
Imports System.EnterpriseServices

Public Class ClsCriteria
    'Classe responsável por montar uma expressão SQL a partir 
    'do objeto passado no construtor e dos critérios passados a partir
    'de nomes de propriedades, valores e critérios de seleção

#Region "Declarações de variáveis"
    Private ClConfig As ClsConfig 'Classe de configuração da camada de persistência
    Private lTable As String      'Tabela do banco de dados 
    Private lSQLWhere As String    'Clausula where montada através dos métodos Add
    Private lType As Type                           'Recebe o type do objeto através do método GETType()
    Private lAssembly As System.Reflection.Assembly '
    Private lField() As ClsSQL.stField              'Coleção de fields representando o objeto e relacionando com os campos na tabela
    Private lCollSell As ArrayList 'Coleção de objetos de seleção
    Private lSQLCriteria As String 'Expressão SQL base gerada
    Private lObjectBase As Object  'Objeto base do critério gerado
#End Region

#Region "Enumeradores"
    Public Enum CriteriaType 'Tipos de critério de seleção
        Equal = 0
        NotEqual = 1
        Greater = 3
        GreaterOrEqual = 4
        Less = 5
        LessOrEqual = 6
    End Enum

    Public Enum LikeCriteria 'Tipos de like possíveis
        isBegin = 0
        isAnyPlace = 1
        isEnd = 2
    End Enum

    Public Enum Union 'Como os critérios irão interagir entre si
        uOr = 0
        uAnd = 1
    End Enum
#End Region


    Public Sub New(ByVal ObjectBase As Object, ByVal XMLPath As String)
        'Construtor recebe o objeto base da seleção
        lObjectBase = ObjectBase

        'Instancia a classe de configuração
        ClConfig = New ClsConfig(XMLPath)
        lTable = ClConfig.DataBaseTableName(ObjectBase)

        '*************************************************************************************
        'Pega as propriedades do objeto passado
        lType = ObjectBase.GetType
        lAssembly = lType.Assembly
        lField = ClConfig.FieldsCharge(ObjectBase)
    End Sub

    Public Sub New(ByVal ObjectBase As Object, ByVal oXML As Xml.XmlDocument)
        'Construtor recebe o objeto base da seleção
        lObjectBase = ObjectBase

        'Instancia a classe de configuração
        ClConfig = New ClsConfig(oXML)
        lTable = ClConfig.DataBaseTableName(ObjectBase)

        '*************************************************************************************
        'Pega as propriedades do objeto passado
        lType = ObjectBase.GetType
        lAssembly = lType.Assembly
        lField = ClConfig.FieldsCharge(ObjectBase)
    End Sub

    Public Property SQLWhere() As String
        Get
            Return lSQLWhere
        End Get
        Set(ByVal Value As String)
            lSQLWhere = Value
        End Set
    End Property

    Private Function GetConector(ByVal pUnion As Union) As String
        'Retorna o conector que será utilizado na montagem do critéiro
        Select Case pUnion
            Case Union.uAnd
                Return " and "
            Case Union.uOr
                Return " or "
            Case Else
                Return ""
        End Select
    End Function

    Private Function GetExpr(ByVal pCriteriaType As CriteriaType) As String
        'Retorna o operador de seleção que será utilizado na expressão de seleção
        Select Case pCriteriaType
            Case CriteriaType.Equal
                Return "="
            Case CriteriaType.Greater
                Return ">"
            Case CriteriaType.GreaterOrEqual
                Return ">="
            Case CriteriaType.Less
                Return "<"
            Case CriteriaType.LessOrEqual
                Return "<="
            Case CriteriaType.NotEqual
                Return "<>"
            Case Else
                Return ""
        End Select
    End Function

    Private Function MontLike(ByVal pLike As LikeCriteria, ByVal pField As ClsSQL.stField, ByVal pValue As Object, ByVal pNotLike As Boolean) As String
        'Monta o valor do like, conforme banco de dados e critério desejado
        Dim Text As New Text.StringBuilder
        Dim Special As String
        Select Case ClConfig.DataBaseType
            Case ClsSQL.EnumDBType.OLEDB
                Special = "*"
            Case Else
                Special = "%"
        End Select
        Text.Append(pField.TableFieldName)
        Select Case pField.Type
            Case ClsSQL.EnumType.eBin
                Err.Raise(vbObjectError + 17001, Me.ToString, "The binary colluns can't used in where expression !")
            Case ClsSQL.EnumType.eBool
                Text.Append("=")
                Text.Append(Boolean.Parse(pValue))
            Case ClsSQL.EnumType.eDate
                'Formata a data, de acordo com o tipo de banco de dados
                Text.Append("=")
                Select Case ClConfig.DataBaseType
                    Case ClsSQL.EnumDBType.OLEDB
                        Text.Append(Format(Date.Parse(pValue), "yyyy/MM/dd"))
                    Case ClsSQL.EnumDBType.SQLServer, ClsSQL.EnumDBType.Oracle
                        Text.Append(Format(Date.Parse(pValue), "MM/dd/yyyy"))
                    Case Else
                        Text.Append(Format(Date.Parse(pValue), "MM/dd/yyyy"))
                End Select
            Case ClsSQL.EnumType.eDateTime
                Text.Append("=")
                Text.Append("'")
                Select Case ClConfig.DataBaseType
                    Case ClsSQL.EnumDBType.OLEDB
                        Text.Append(Format(DateTime.Parse(pValue), "yyyy-MM-dd HH:mm:ss"))
                    Case ClsSQL.EnumDBType.SQLServer, ClsSQL.EnumDBType.Oracle
                        Text.Append(Format(DateTime.Parse(pValue), "MM-dd-yyyy HH:mm:ss"))
                End Select
                Text.Append("'")
            Case ClsSQL.EnumType.eTime
                Text.Append("=")
                Text.Append("'")
                Text.Append(Format(Date.Parse(pValue), "HH:mm:ss"))
                Text.Append("'")
            Case ClsSQL.EnumType.eFloat
                Text.Append("=")
                Text.Append(Double.Parse(pValue).ToString.Replace(",", "."))
            Case ClsSQL.EnumType.eInteger
                Text.Append("=")
                Text.Append(Int64.Parse(pValue).ToString)
            Case ClsSQL.EnumType.eText
                If pNotLike Then
                    Text.Append(" not like '")
                Else
                    Text.Append(" like '")
                End If

                Select Case pLike
                    Case LikeCriteria.isBegin
                        Text.Append(pValue.ToString)
                        Text.Append(Special)
                    Case LikeCriteria.isEnd
                        Text.Append(Special)
                        Text.Append(pValue.ToString)
                    Case LikeCriteria.isAnyPlace
                        Text.Append(Special)
                        Text.Append(pValue.ToString)
                        Text.Append(Special)
                End Select
                Text.Append("'")
        End Select
        Return Text.ToString
    End Function

    Private Function MontCriteria(ByVal pField As ClsSQL.stField, ByVal pValue As Object, ByVal pCriteriaType As CriteriaType) As String
        'Monta o valor do critério, conforme banco de dados e conector desejado
        Dim Text As New Text.StringBuilder
        Text.Append(pField.TableFieldName)
        Text.Append(GetExpr(pCriteriaType))
        Select Case pField.Type
            Case ClsSQL.EnumType.eBin
                Throw New Exception("The binary colluns can't used in where expression !")
            Case ClsSQL.EnumType.eBool
                If TypeOf (pValue) Is Boolean Then
                    If Boolean.Parse(pValue) = True Then
                        Text.Append("1")
                    Else
                        Text.Append("0")
                    End If
                Else
                    Text.Append(pValue)
                End If
            Case ClsSQL.EnumType.eDate
                Text.Append("'")
                'Formata a data, de acordo com o tipo de banco de dados
                Select Case ClConfig.DataBaseType
                    Case ClsSQL.EnumDBType.OLEDB
                        Text.Append(Format(Date.Parse(pValue), "yyyy/MM/dd"))
                    Case ClsSQL.EnumDBType.SQLServer, ClsSQL.EnumDBType.Oracle
                        Text.Append(Format(Date.Parse(pValue), "MM/dd/yyyy"))
                    Case Else
                        Text.Append(Format(Date.Parse(pValue), "MM/dd/yyyy"))
                End Select
                Text.Append("'")
            Case ClsSQL.EnumType.eDateTime
                Text.Append("'")
                Select Case ClConfig.DataBaseType
                    Case ClsSQL.EnumDBType.OLEDB
                        Text.Append(Format(DateTime.Parse(pValue), "yyyy-MM-dd HH:mm:ss"))
                    Case ClsSQL.EnumDBType.SQLServer, ClsSQL.EnumDBType.Oracle
                        Text.Append(Format(DateTime.Parse(pValue), "MM-dd-yyyy HH:mm:ss"))
                End Select
                Text.Append("'")
            Case ClsSQL.EnumType.eTime
                Text.Append("'")
                Text.Append(Format(Date.Parse(pValue), "HH:mm:ss"))
                Text.Append("'")
            Case ClsSQL.EnumType.eFloat
                Text.Append(Double.Parse(pValue).ToString.Replace(",", "."))
            Case ClsSQL.EnumType.eInteger
                Text.Append(Int64.Parse(pValue).ToString)
            Case ClsSQL.EnumType.eText
                Text.Append("'")
                Text.Append(pValue.ToString)
                Text.Append("'")
        End Select
        Return Text.ToString
    End Function

    Private Function MontSubstringCriteria(ByVal pField As ClsSQL.stField, ByVal pStart As Integer, ByVal pLength As Integer, ByVal pValue As Object, ByVal pCriteriaType As CriteriaType, Optional ByVal strLen As Integer = 13) As String
        'Monta o valor do critério, conforme banco de dados e conector desejado
        Dim Text As New Text.StringBuilder
        Text.Append(" Substring(")
        If pField.Type = ClsSQL.EnumType.eText Then
            Text.Append(pField.TableFieldName)
        ElseIf pField.Type = ClsSQL.EnumType.eFloat Or pField.Type = ClsSQL.EnumType.eInteger Then
            Text.Append("str(")
            Text.Append(pField.TableFieldName)
            Text.Append("," & strLen & ")")
        Else
            Throw New Exception("The type can't used in substring expression!")
        End If
        Text.Append(",")
        Text.Append(pStart)
        Text.Append(",")
        Text.Append(pLength)
        Text.Append(")")
        Text.Append(GetExpr(pCriteriaType))
        Text.Append("'")
        Text.Append(pValue.ToString)
        Text.Append("'")
        Return Text.ToString
    End Function

    Private Function MontBetween(ByVal pField As ClsSQL.stField, ByVal PBeginValue As Object, ByVal pEndValue As Object)
        'Monta o valor do critério, conforme banco de dados e conector desejado
        Dim Text As New Text.StringBuilder
        Dim i As Byte, pValue As Object
        For i = 0 To 1
            If i = 0 Then
                pValue = PBeginValue
                Text.Append(pField.TableFieldName)
                Text.Append(" Between ")
            Else
                pValue = pEndValue
                Text.Append(" And ")
            End If
            Select Case pField.Type
                Case ClsSQL.EnumType.eBin
                    Err.Raise(vbObjectError + 17001, Me, "The binary colluns can't used in where expression !")
                Case ClsSQL.EnumType.eBool
                    Text.Append(Boolean.Parse(pValue))
                Case ClsSQL.EnumType.eDate
                    'Formata a data, de acordo com o tipo de banco de dados
                    Text.Append("'")
                    Select Case ClConfig.DataBaseType
                        Case ClsSQL.EnumDBType.OLEDB

                            Text.Append(Format(Date.Parse(pValue), "yyyy/MM/dd"))

                        Case ClsSQL.EnumDBType.SQLServer, ClsSQL.EnumDBType.Oracle

                            Text.Append(Format(Date.Parse(pValue), "MM/dd/yyyy"))

                        Case Else
                            Text.Append(Format(Date.Parse(pValue), "MM/dd/yyyy"))
                    End Select
                    Text.Append("'")
                Case ClsSQL.EnumType.eDateTime
                    Text.Append("'")
                    Select Case ClConfig.DataBaseType
                        Case ClsSQL.EnumDBType.OLEDB
                            Text.Append(Format(DateTime.Parse(pValue), "yyyy-MM-dd HH:mm:ss"))
                        Case ClsSQL.EnumDBType.SQLServer, ClsSQL.EnumDBType.Oracle
                            Text.Append(Format(DateTime.Parse(pValue), "MM-dd-yyyy HH:mm:ss"))
                    End Select
                    Text.Append("'")
                Case ClsSQL.EnumType.eTime
                    Text.Append("'")
                    Text.Append(Format(Date.Parse(pValue), "HH:mm:ss"))
                    Text.Append("'")
                Case ClsSQL.EnumType.eFloat
                    Text.Append(Double.Parse(pValue).ToString.Replace(",", "."))
                Case ClsSQL.EnumType.eInteger
                    Text.Append(Int64.Parse(pValue).ToString)
                Case ClsSQL.EnumType.eText
                    Text.Append("'")
                    Text.Append(pValue.ToString)
                    Text.Append("'")
            End Select
        Next
        Return Text.ToString
    End Function

    Private Function MontInList(ByVal pField As ClsSQL.stField, ByVal pValues As ArrayList) As String
        'monta a lista do in, de acordo com o tipo do campo
        'O for está dentro do critério, para que não perca tempo com critério a cada elemento do in
        'pField = Campo do critério
        'PValues -> Array de valores utilizados para montar o in ou o not in
        Dim i As Integer
        Dim lReturn As New System.Text.StringBuilder
        Select Case pField.Type
            Case ClsSQL.EnumType.eBin
                Err.Raise(vbObjectError + 17001, Me, "The binary colluns can't used in where expression !")
            Case ClsSQL.EnumType.eText
                For i = 0 To pValues.Count - 1
                    lReturn.Append("'")
                    lReturn.Append(pValues(i))
                    lReturn.Append("'")
                    If i < pValues.Count - 1 Then
                        lReturn.Append(",")
                    End If
                Next
            Case ClsSQL.EnumType.eTime
                For i = 0 To pValues.Count - 1
                    lReturn.Append("'")
                    lReturn.Append(Format(Date.Parse(pValues(i).ToString), "hh:mm:ss"))
                    lReturn.Append("'")
                    If i < pValues.Count - 1 Then
                        lReturn.Append(",")
                    End If
                Next

            Case ClsSQL.EnumType.eDate
                lReturn.Append("'")
                For i = 0 To pValues.Count - 1
                    Select Case ClConfig.DataBaseType
                        Case ClsSQL.EnumDBType.OLEDB
                            lReturn.Append(Format(Date.Parse(pValues(i).ToString), "yyyy/MM/dd"))
                        Case ClsSQL.EnumDBType.SQLServer, ClsSQL.EnumDBType.Oracle
                            lReturn.Append(Format(Date.Parse(pValues(i).ToString), "MM/dd/yyyy"))
                        Case Else
                            lReturn.Append(Format(Date.Parse(pValues(i).ToString), "MM/dd/yyyy"))
                    End Select
                    If i < pValues.Count - 1 Then
                        lReturn.Append(",")
                    End If
                Next
                lReturn.Append("'")
            Case ClsSQL.EnumType.eDateTime
                For i = 0 To pValues.Count - 1
                    lReturn.Append("'")
                    Select Case ClConfig.DataBaseType
                        Case ClsSQL.EnumDBType.OLEDB
                            lReturn.Append(Format(DateTime.Parse(pValues(i).ToString), "yyyy-MM-dd HH:mm:ss"))
                        Case ClsSQL.EnumDBType.SQLServer, ClsSQL.EnumDBType.Oracle
                            lReturn.Append(Format(DateTime.Parse(pValues(i).ToString), "MM-dd-yyyy HH:mm:ss"))
                    End Select
                    lReturn.Append("'")
                    If i < pValues.Count - 1 Then
                        lReturn.Append(",")
                    End If
                Next
            Case ClsSQL.EnumType.eFloat
                For i = 0 To pValues.Count - 1
                    lReturn.Append(Format(pValues(i).ToString.Replace(",", ".")))
                    If i < pValues.Count - 1 Then
                        lReturn.Append(",")
                    End If
                Next
            Case ClsSQL.EnumType.eInteger
                Dim Value As String
                For i = 0 To pValues.Count - 1
                    Value = pValues(i).ToString.Replace(",", "")
                    Value = pValues(i).ToString.Replace(".", "")
                    lReturn.Append(Value)
                    If i < pValues.Count - 1 Then
                        lReturn.Append(",")
                    End If
                Next
            Case ClsSQL.EnumType.eBool
                Err.Raise(vbObjectError + 17015, Me.ToString, "The boolean colluns can't used in IN expression !")
        End Select
        Return lReturn.ToString
    End Function

    Public Function AddCriteria(ByVal pRow As String, ByVal pValue As Object, ByVal pCriteriaType As CriteriaType, Optional ByVal pUnion As Union = Union.uAnd) As String
        'Adiciona critério a expressão de seleção montada
        'Critérios normais, com =, >,>=,<, <= e <>
        '-------------PARAMETROS RECEBIDOS-------------
        'pRow          -> Objeto.NomeDaPropriedade que irá participar do critério de seleção
        'pValue        -> Valor que pertencerá ao critério de seleção
        'pCriteriaType -> Tipo de critério que será passado
        'pUnion        -> Como o critério será concatenado com o critério anterior
        '---------FIM DOS PARAMETROS RECEBIDOS-------------
        Dim lSQL As New System.Text.StringBuilder
        Dim i As Byte, Find As Boolean
        For i = 0 To lField.Length - 1
            If lField(i).ObjFieldName.ToLower = pRow.ToLower Then
                If lSQLWhere = "" Then
                    lSQL.Append(" where ")
                Else
                    lSQL.Append(GetConector(pUnion))
                End If
                lSQL.Append(MontCriteria(lField(i), pValue, pCriteriaType))
                Find = True
                Exit For
            End If
        Next
        If Find = False Then
            Err.Raise(vbObjectError + 17000, Me.ToString, "The property value passed in pRow parameter is not correct!")
        End If
        'Acrescenta o texto montado no lSQLWhere e o retorna na função
        lSQLWhere += lSQL.ToString
        Return lSQLWhere
    End Function

    Public Function AddLikeCriteria(ByVal pRow As String, ByVal pValue As Object, ByVal pLike As LikeCriteria, Optional ByVal pUnion As Union = Union.uAnd, Optional ByVal pNotLike As Boolean = False) As String
        'Adiciona critério a expressão de seleção montada
        'Somente para critérios que utilizem like
        '-------------PARAMETROS RECEBIDOS-------------
        'pRow    -> Objeto.NomeDaPropriedade que irá participar do critério de seleção
        'pValue  -> Valor que pertencerá ao critério de seleção
        'pLike   -> Se o critério utilizando o like utilizará caracter coringa no inicio, no final ou no inicio e no final
        'pUnion  -> Como o critério será concatenado com o critério anterior
        'pNotLike -> Define que o critério será not like campo.propriedade not like
        '---------FIM DOS PARAMETROS RECEBIDOS-------------      
        Dim lSQL As New System.Text.StringBuilder
        Dim i As Byte, Find As Boolean
        For i = 0 To lField.Length - 1
            If lField(i).ObjFieldName.ToLower = pRow.ToLower Then
                If lSQLWhere = "" Then
                    lSQL.Append(" where ")
                Else
                    lSQL.Append(GetConector(pUnion))
                End If
                lSQL.Append(MontLike(pLike, lField(i), pValue, pNotLike))
                Find = True
                Exit For
            End If
        Next
        If Find = False Then
            Err.Raise(vbObjectError + 17000, Me.ToString, "The property value passed in pRow parameter is not correct!")
        End If
        'Acrescenta o texto montado no lSQLWhere e o retorna na função
        lSQLWhere += lSQL.ToString
        Return lSQLWhere
    End Function

    Public Function AddBetweenCriteria(ByVal pRow As String, ByVal pBeginValue As Object, ByVal pEndValue As Object, Optional ByVal pUnion As Union = Union.uAnd) As String
        'Adiciona critério a expressão de seleção montada
        'Expressão Between
        '-------------PARAMETROS RECEBIDOS-------------
        'pRow        -> Objeto.NomeDaPropriedade que irá participar do critério de seleção
        'pBeginValue -> Valor inicial do between
        'pEndValue   -> Valor final do between 
        'pUnion      -> Como o critério será concatenado com o critério anterior
        '---------FIM DOS PARAMETROS RECEBIDOS-------------
        Dim lSQL As New System.Text.StringBuilder
        Dim i As Byte, Find As Boolean
        For i = 0 To lField.Length - 1
            If lField(i).ObjFieldName.ToLower = pRow.ToLower Then
                If lSQLWhere = "" Then
                    lSQL.Append(" where ")
                Else
                    lSQL.Append(GetConector(pUnion))
                End If
                lSQL.Append(MontBetween(lField(i), pBeginValue, pEndValue))
                Find = True
                Exit For
            End If
        Next
        If Find = False Then
            Err.Raise(vbObjectError + 17000, Me.ToString, "The property value passed in pRow parameter is not correct!")
        End If
        'Acrescenta o texto montado no lSQLWhere e o retorna na função
        lSQLWhere += lSQL.ToString
        Return lSQLWhere
    End Function

    Public Function AddSelectInCriteria(ByVal pRow As String, ByVal pRowSubSelect As String, ByVal ObjSubSelect As Object, Optional ByVal pCriterioSubSelect As ClsCriteria = Nothing, Optional ByVal pUnion As Union = Union.uAnd, Optional ByVal pNotIn As Boolean = False) As String
        'Adiciona critério a expressão de seleção montada
        'Expressão com in ou not in para sub-selects
        '-------------PARAMETROS RECEBIDOS-------------
        'pRow               -> NomeDaPropriedade que irá participar do critério de seleção
        'pRowSubSelect      -> NomeDaPropriedade que irá ser o campo base do sub-select
        'pobjSubSelect      -> Objeto de referência do sub-select
        'pCriterioSubSelect -> Critério do sub-select. 
        'pUnion             -> Como o critério será concatenado com o critério anterior
        'pNotIn             -> Determina se o operador not será utilizado antes do in
        '---------FIM DOS PARAMETROS RECEBIDOS-------------
        Dim lSQL As New System.Text.StringBuilder
        Dim i As Byte, Find As Boolean
        Dim TableName As String
        Dim lSubField() As ClsSQL.stField
        Dim SubRowName As String = ""
        lSubField = ClConfig.FieldsCharge(ObjSubSelect) 'Carrega a coleção dos campos da tabela
        TableName = ClConfig.DataBaseTableName(ObjSubSelect) 'Busca o nome da tabela no banco de dados

        'Busca o nome do campo que será parte do critério do subselect, dentro do in ou not in
        For i = 0 To lSubField.Length - 1
            If lSubField(i).ObjFieldName.ToLower = pRowSubSelect.ToLower Then
                SubRowName = lSubField(i).TableFieldName
                Find = True
                Exit For
            End If
        Next
        If Find = False Then
            Err.Raise(vbObjectError + 17000, Me.ToString, "The property value passed in pRowSubSelect parameter is not correct!")
        End If

        Find = False
        'A cláusula where é montada dentro do loop, quando o registro do critério é encontrado
        For i = 0 To lField.Length - 1
            If lField(i).ObjFieldName.ToLower = pRow.ToLower Then
                If lSQLWhere = "" Then
                    lSQL.Append(" where ")
                Else
                    lSQL.Append(GetConector(pUnion))
                End If
                lSQL.Append(lField(i).TableFieldName)
                If pNotIn = True Then
                    lSQL.Append(" not ")
                End If
                lSQL.Append(" in (Select ")
                lSQL.Append(SubRowName)
                lSQL.Append(" from ")
                lSQL.Append(TableName)
                'Verifica a clausula where, caso o objeto não seja igual a nothing
                If Not IsNothing(pCriterioSubSelect) Then
                    lSQL.Append(pCriterioSubSelect.SQLWhere)
                End If
                lSQL.Append(")")
                Find = True
                Exit For
            End If
        Next
        If Find = False Then
            Err.Raise(vbObjectError + 17000, Me.ToString, "The property value passed in pRow parameter is not correct!")
        End If
        'Acrescenta o texto montado no lSQLWhere e o retorna na função
        lSQLWhere += lSQL.ToString
        Return lSQLWhere
    End Function

    Public Function AddListInCriteria(ByVal pRow As String, ByVal pValues As ArrayList, Optional ByVal pUnion As Union = Union.uAnd, Optional ByVal pNotIn As Boolean = False) As String
        'Adiciona critério a expressão de seleção montada
        'Expressão com in ou not in para array de valores
        '-------------PARAMETROS RECEBIDOS-------------
        'pRow    -> Objeto.NomeDaPropriedade que irá participar do critério de seleção
        'PValues -> Array de valores utilizados para montar o in ou o not in
        'pUnion  -> Como o critério será concatenado com o critério anterior
        'pNotIn  -> Determina se o operador not será utilizado antes do in
        '---------FIM DOS PARAMETROS RECEBIDOS-------------
        Dim lSQL As New System.Text.StringBuilder
        Dim i As Byte, Find As Boolean
        Find = False
        'A cláusula where é montada dentro do loop, quando o registro do critério é encontrado
        For i = 0 To lField.Length - 1
            If lField(i).ObjFieldName.ToLower = pRow.ToLower Then
                If lSQLWhere = "" Then
                    lSQL.Append(" where ")
                Else
                    lSQL.Append(GetConector(pUnion))
                End If
                lSQL.Append(lField(i).TableFieldName)
                If pNotIn = True Then
                    lSQL.Append(" not ")
                End If
                lSQL.Append(" in (")

                lSQL.Append(MontInList(lField(i), pValues)) 'Monta a lista de parâmetros do IN

                lSQL.Append(")")
                Find = True
                Exit For
            End If
        Next
        If Find = False Then
            Throw New Exception("The property value passed in pRow parameter is not correct!")
        End If
        'Acrescenta o texto montado no lSQLWhere e o retorna na função
        lSQLWhere += lSQL.ToString
        Return lSQLWhere
    End Function

    Public Function AddSubstringCriteria(ByVal pRow As String, ByVal pValue As Object, ByVal pStart As Integer, ByVal pLength As Integer, Optional ByVal strLen As Integer = 13, Optional ByVal pCriteriaType As CriteriaType = CriteriaType.Equal, Optional ByVal pUnion As Union = Union.uAnd) As String
        'Adiciona critério a expressão de seleção montada
        'Expressão com in ou not in para array de valores
        '-------------PARAMETROS RECEBIDOS-------------
        'pRow    -> Objeto.NomeDaPropriedade que irá participar do critério de seleção
        'PValue  -> valor utilizados para montar o substring
        'pStart  -> Início da expressão substring
        'pLength -> Final da expressão substring
        'pUnion  -> Como o critério será concatenado com o critério anterior

        '---------FIM DOS PARAMETROS RECEBIDOS-------------
        Dim lSQL As New System.Text.StringBuilder
        Dim i As Byte, Find As Boolean
        Find = False
        'A cláusula where é montada dentro do loop, quando o registro do critério é encontrado
        For i = 0 To lField.Length - 1
            If lField(i).ObjFieldName.ToLower = pRow.ToLower Then
                If lSQLWhere = "" Then
                    lSQL.Append(" where ")
                Else
                    lSQL.Append(GetConector(pUnion))
                End If
                lSQL.Append(MontSubstringCriteria(lField(i), pStart, pLength, pValue, pCriteriaType, strLen))
                'lSQL.Append(MontInList(lField(i), pValues)) 'Monta a lista de parâmetros do IN
                'lSQL.Append(")")
                Find = True
                Exit For
            End If
        Next
        If Find = False Then
            Throw New Exception("The property value passed in pRow parameter is not correct!")
        End If
        'Acrescenta o texto montado no lSQLWhere e o retorna na função
        lSQLWhere += lSQL.ToString
        Return lSQLWhere
    End Function

    Public Function AddNullCriteria(ByVal pRow As String, Optional ByVal NotNull As Boolean = False, Optional ByVal pUnion As Union = Union.uAnd)
        'Adiciona critério a expressão de seleção montada
        'Expressão com is null ou is not null 
        '-------------PARAMETROS RECEBIDOS-------------
        'pRow    -> Objeto.NomeDaPropriedade que irá participar do critério de seleção
        'NotNull -> Define se o critério aplicado será o is null ou o is not null
        'pUnion  -> Como o critério será concatenado com o critério anterior
        '---------FIM DOS PARAMETROS RECEBIDOS-------------
        Dim lSQL As New System.Text.StringBuilder
        Dim i As Byte, Find As Boolean
        For i = 0 To lField.Length - 1
            If lField(i).ObjFieldName.ToLower = pRow.ToLower Then
                If lSQLWhere = "" Then
                    lSQL.Append(" where ")
                Else
                    lSQL.Append(GetConector(pUnion))
                End If
                lSQL.Append(lField(i).TableFieldName)
                If NotNull Then
                    lSQL.Append(" is not null ")
                Else
                    lSQL.Append(" is null ")
                End If
                Find = True
                Exit For
            End If
        Next
        If Find = False Then
            Throw New Exception("The property value passed in pRow parameter is not correct!")
        End If
        'Acrescenta o texto montado no lSQLWhere e o retorna na função
        lSQLWhere += lSQL.ToString
        Return lSQLWhere
    End Function

    Public Sub ClearCriteria()
        'Limpa o critério de seleção
        lSQLWhere = ""
    End Sub

    Public Sub UnionCriteria(ByVal pCriteria As ClsCriteria, ByVal pUnion As Union)
        'Une dois critérios (o existente atualmente na classe, com o que está dentro do
        'objeto pCriteria passado, separando-os por parênteses.
        'Para possibilitar a crição de critérios do tipo (Campo1=valor and campo1=valor1) and (campo2=valor2 or campo2=valor3)
        '-------------PARAMETROS RECEBIDOS-------------
        'pCriteria -> Objeto criteria contendo um critério para ser       concatenado()
        'pUnion    -> Como o critério será concatenado com o critério anterior
        '---------FIM DOS PARAMETROS RECEBIDOS-------------
        Dim lSQL As New Text.StringBuilder("(")
        lSQL.Append(lSQLWhere.Replace("where", "").Trim)
        lSQL.Append(") ")
        lSQL.Append(GetConector(pUnion))
        lSQL.Append(" (")
        lSQL.Append(pCriteria.SQLWhere.Substring(6, pCriteria.SQLWhere.Length - 6).Trim)
        lSQL.Append(")")
        lSQL.Insert(0, " where ")
        lSQLWhere = lSQL.ToString
    End Sub

    Protected Overrides Sub Finalize()
        ClConfig = Nothing
        MyBase.Finalize()
    End Sub

End Class
