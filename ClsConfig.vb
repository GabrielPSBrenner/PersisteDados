Imports System.Reflection
Imports System.EnterpriseServices

Public Class ClsConfig
   'Lê e retorna para as demais classes o XML de configuração do site

   Private lXMLDoc As Xml.XmlDocument          'Variável interna representando o objeto XML que representa o arquivo
   Private lDataBaseType As ClsSQL.EnumDBType  'Define qual banco de dados está sendo acessado
   Private lConnectionString As String         'String de conexão do banco de dados

#Region "Mount Relations"
   '*********************************************************************************************************************
   '*********************************************************************************************************************
   Public Structure stRelations
      'Estrutura criada com o objetivo de ajudar na montagem dos relacionamentos
      Public FromTable As String          'Tabela de origem 
      Public ToTable As String            'Tabela de referência do relacionamento
      Public objName As String            'Nome do objeto que será instanciado
      Public objPropertyName As String    'Nome da propriedade no objeto
      Public Multiplicity As String       'Define a multiplicidade (* = Tabela filho, 1 = Tabela relacionada através de uma FK)
      Public RowFromTable() As String     'Nome da(s) coluna(s) de origem na tabela
      Public RowToTable() As String       'Nome da(s) coluna(s) de destino na tabela
   End Structure

   Public Structure StRelationsArray
      Public Item() As stRelations
   End Structure

   Public Function MontRelations(ByVal objFrom As Object, Optional ByVal pTableName As String = "") As StRelationsArray
      'ObjNodeList = parte do XML que representa o objeto passado
      '*********************************************************************************
      Dim i, j As Byte
      Dim TableName As String
      Dim ObjName As String = objFrom.GetType.Name.ToString 'Busca o nome do objeto
      Dim ObjRelations() As stRelations

      Dim ObjNodeList As Xml.XmlNodeList = lXMLDoc.GetElementsByTagName(ObjName)
      '********************************************************************************************
      'Caso o nome da tabela de origem não seja passada como parâmetro, o rotina solicita o mesmo
      If pTableName = "" Then
         pTableName = ObjNodeList.Item(0).Attributes("Table").InnerText
      Else
         TableName = pTableName
      End If
      '********************************************************************************************
      'Busca as relações que o objeto pode conter

      ObjNodeList = ObjNodeList.Item(0).SelectNodes("Relations")

      If ObjNodeList.Count > 0 Then

         '-----------------------------------------------------------------------------------------------------
         'Monta um objeto contendo os relacionamentos identificados no XML
         '-----------------------------------------------------------------------------------------------------
         ReDim ObjRelations(ObjNodeList.Item(0).ChildNodes.Count - 1)
         For i = 0 To ObjNodeList.Item(0).ChildNodes.Count - 1
            With ObjRelations(i)
               .FromTable = pTableName
               .ToTable = ObjNodeList.Item(0).ChildNodes(i).Attributes("TableName").InnerText
               .objPropertyName = ObjNodeList.Item(0).ChildNodes(i).Attributes("ObjPropertyName").InnerText
               .Multiplicity = ObjNodeList.Item(0).ChildNodes(i).Attributes("Multiplicity").InnerText
               .objName = ObjNodeList.Item(0).ChildNodes(i).Attributes("objName").InnerText

               Dim objListFields As Xml.XmlNodeList = ObjNodeList.Item(0).ChildNodes(i).SelectNodes("RowLink")


               ReDim .RowFromTable(objListFields.Count - 1)
               ReDim .RowToTable(objListFields.Count - 1)
               'Monta a coleção de fields que se relacionam
               For j = 0 To objListFields.Count - 1
                  .RowFromTable(j) = objListFields.Item(j).Attributes("FromTableRowName").InnerText
                  .RowToTable(j) = objListFields.Item(j).Attributes("ToTableRowName").InnerText
               Next
            End With
         Next
         '-----------------------------------------------------------------------------------------------------
      End If
      'Return ObjRelations
   End Function

   '*********************************************************************************************************************
   '*********************************************************************************************************************
#End Region

   Public Sub New(ByVal oXML As Xml.XmlDocument)
      Try
         lXMLDoc = oXML
         StartConfig() 'Ativa a configuração inicial
      Catch ex As Exception
         Throw New Exception(Err.Description, ex)
      End Try
   End Sub

   Public Sub New(ByVal XMLPath As String)
      'Recebe o path do XML de configuração
      Try
         lXMLDoc = New Xml.XmlDocument
         lXMLDoc.Load(XMLPath)
         StartConfig() 'Ativa a configuração inicial
      Catch ex As Xml.XmlException
         'Tratar o erro aqui
         Throw New Exception(Err.Description, ex)
      End Try
   End Sub

   Private Sub StartConfig()
      Dim objNodeList As Xml.XmlNodeList
      '**************************************************************
      objNodeList = lXMLDoc.GetElementsByTagName("ServerType")
      Select Case objNodeList.Item(0).InnerText.ToString.ToLower
         Case "sqlserver", "sql-server"
            lDataBaseType = ClsSQL.EnumDBType.SQLServer
         Case "access", "oledb"
            lDataBaseType = ClsSQL.EnumDBType.OLEDB
         Case "mysql", "my-sql"
            lDataBaseType = ClsSQL.EnumDBType.MySQL
         Case "oracle"
            lDataBaseType = ClsSQL.EnumDBType.Oracle
         Case "odbc"
            lDataBaseType = ClsSQL.EnumDBType.ODBC
      End Select
      objNodeList = lXMLDoc.GetElementsByTagName("ConnectionString") 'Busca a string de conexão que será utilizada pela camada
      lConnectionString = objNodeList.Item(0).InnerText.ToString
      '**************************************************************
   End Sub

   Public Function GetProcedureDetails(ByVal ObjName As String) As Xml.XmlNode
      Dim noRetorno As Xml.XmlNode = lXMLDoc.SelectSingleNode("//Objects/Procedures/Procedure[@object_name='" & ObjName & "']")
      Return noRetorno
   End Function

   Public ReadOnly Property DataBaseType() As ClsSQL.EnumDBType 'Retorna o tipo de banco de dados
      Get
         Return lDataBaseType
      End Get
   End Property

   Public ReadOnly Property ConnectionString() As String
      Get
         Return lConnectionString
      End Get
   End Property

   Public Function DataBaseTableName(ByVal ObjRef As Object, Optional ByRef Prefix As String = "") As String
      'Busca o nome da tabela no banco de dados, lendo o XML de configuração, a partir de um objeto
      'ObjRef = Objeto instanciado
      Dim lType As Type
      lType = ObjRef.GetType
      Return DataBaseTableName(lType.Name.ToString, Prefix)
   End Function

   Public Function DataBaseTableName(ByVal ObjName As String, Optional ByRef Prefix As String = "") As String
      'Busca o nome da tabela no banco de dados, lendo o XML de configuração, a partir do nome do objeto
      'ObjName = Nome do objeto instanciado
      Dim objNodeList As Xml.XmlNodeList
      objNodeList = lXMLDoc.GetElementsByTagName(ObjName)
      If objNodeList.Count = 0 Then
         Err.Raise(vbObjectError + 15001, Me, "The XML attributes is not correct or the relation with this object not exist! ")
      End If
      Try
         Prefix = objNodeList.Item(0).Attributes("PrefixTableName").Value
      Catch ex As Exception
         Prefix = ""
      End Try
      Return objNodeList.Item(0).Attributes("Table").Value
   End Function

   Public Function GetFields(ByVal objRef As Object, Optional ByRef TableName As String = "") As Xml.XmlNodeList
      'Lê o XML e retorna um nodelist contendo toda a coleção de fields que contém a relação do objeto com o XML
      'objRef = Objeto de referência
      Dim lType As Type
      lType = objRef.GetType
      Return GetFields(lType.Name.ToString, TableName)
   End Function

   Public Function GetFields(ByVal objName As String, ByRef TableName As String) As Xml.XmlNodeList
      'Lê o XML e retorna um nodelist contendo toda a coleção de fields que contém a relação do objecto com o XML
      'ObjName = Nome do objeto
      Dim objNodeList As Xml.XmlNodeList
      objNodeList = lXMLDoc.GetElementsByTagName(objName)
      TableName = objNodeList.Item(0).Attributes("Table").InnerText 'Busca o nome da tabela      
      objNodeList = objNodeList.Item(0).SelectNodes("Fields")
      Return objNodeList
   End Function

   Public Function GetFieldName(ByVal ObjName As String, ByVal PropertyName As String) As String
      Dim objNodeList As Xml.XmlNodeList
      Dim Response As String
      objNodeList = lXMLDoc.GetElementsByTagName(ObjName)
      objNodeList = objNodeList.Item(0).SelectNodes("Fields")
      objNodeList = objNodeList.Item(0).SelectNodes(PropertyName)
      If objNodeList.Count > 0 Then
         Response = objNodeList.Item(0).InnerText
      Else
         Response = ""
      End If
      objNodeList = Nothing
      Return Response
   End Function

   '''Public Function FieldCharge(ByVal ObjName As String, ByVal PropertyName As String) As ClsSQL.stField
   '''   Dim objNodeList As Xml.XmlNodeList
   '''   objNodeList = lXMLDoc.GetElementsByTagName(ObjName)
   '''   objNodeList = objNodeList.Item(0).SelectNodes("/Fields/" + PropertyName)
   '''   If objNodeList.Count > 0 Then
   '''      Dim Field As New ClsSQL.stField
   '''      Field.ObjFieldName = PropertyName
   '''      Field.TableFieldName = objNodeList.Item(0).InnerText
   '''      ''field.Key ='
   '''   Else
   '''      Return Nothing
   '''   End If
   '''End Function

   Public Function FieldsCharge(ByRef objRef As Object) As ClsSQL.stField()
      Dim Field() As ClsSQL.stField 'Coleção de fields que será retornada
      Dim i As Byte 'Contador utilizado no loop do objeto 
      Dim Pri As PropertyInfo 'Objeto utilizado para ler e atribuir valores ao objeto recebido
      Dim lType As Type 'Objeto que recebe o tipo do objeto recebido
      Dim TypeD As String 'Variável auxiliar
      Dim objNodeList As Xml.XmlNodeList 'Objeto nodelist que recebe a coleção de fields do XML
      Dim TableName As String = ""
      lType = objRef.GetType
      objNodeList = GetFields(lType.Name.ToString, TableName) 'Busca a coleção de fields do XML            
      ReDim Field(objNodeList.Item(0).ChildNodes.Count - 1) 'Redimensiona a coleção de objetos que será retornada
      For i = 0 To objNodeList.Item(0).ChildNodes.Count - 1
         Field(i).TableName = TableName
         Field(i).TableFieldName = objNodeList.Item(0).ChildNodes(i).InnerText 'Get the table field name and mont string
         Field(i).ObjFieldName = objNodeList.Item(0).ChildNodes(i).Name 'Get the obj field name

         Try
            'Get the primary key information
            Field(i).Key = (objNodeList.Item(0).ChildNodes(i).Attributes("Id").Value = "PK")
         Catch
            Field(i).Key = False
         End Try

         Try
            'Get the key information
            Field(i).FK = (objNodeList.Item(0).ChildNodes(i).Attributes("Id").Value = "FK")
         Catch
            Field(i).FK = False
         End Try

         Try
            'Get the autonumera information
            Field(i).AutoNumera = (objNodeList.Item(0).ChildNodes(i).Attributes("AutoNumera").Value = "T")
         Catch
            Field(i).AutoNumera = False
         End Try

         'Verifica o tipo do campo do objeto e da tabela
         TypeD = objNodeList.Item(0).ChildNodes(i).Attributes("Type").Value.ToLower
         Select Case TypeD
            Case "text"
               Field(i).Type = ClsSQL.EnumType.eText
            Case "integer"
               Field(i).Type = ClsSQL.EnumType.eInteger
            Case "float"
               Field(i).Type = ClsSQL.EnumType.eFloat
            Case "date"
               Field(i).Type = ClsSQL.EnumType.eDate
            Case "time"
               Field(i).Type = ClsSQL.EnumType.eTime
            Case "datetime"
               Field(i).Type = ClsSQL.EnumType.eDateTime
            Case "bin"
               Field(i).Type = ClsSQL.EnumType.eBin
            Case "bool"
               Field(i).Type = ClsSQL.EnumType.eBool
         End Select
         Pri = lType.GetProperty(Field(i).ObjFieldName) 'Recupera a propriedade do objeto
         Field(i).Value = Pri.GetValue(objRef, Nothing) 'Lê o valor da propriedade do objeto e atribui ao registro correspondente da coleção
      Next
      Return Field
   End Function

   Protected Overrides Sub Finalize()
      lXMLDoc = Nothing
      MyBase.Finalize()
   End Sub
End Class
