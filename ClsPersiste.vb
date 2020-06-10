Imports System.Reflection
Imports System.EnterpriseServices

<Transaction(TransactionOption.Supported), EventTrackingEnabled(True)> Public Class ClsPersiste
   Implements IDisposable

   Private lStateTrans As EnumStateTrans
   Private lXMLPath As String
   Private ClConfig As ClsConfig

   Public ReadOnly Property StateTrans() As EnumStateTrans
      Get
         Return lStateTrans
      End Get
   End Property

#Region "Enumeradores"
   Public Enum EnumDBType
      SQLServer = 0
      Oracle = 1
      MySQL = 2
      OLEDB = 3
      ODBC = 4
   End Enum

   Public Enum EnumStateTrans
      NotTransaction = 0
      InTransaction = 1
   End Enum
#End Region

#Region "Variáveis de banco de dados"
   Private lDataBase As EnumDBType  'Tipo do banco de dados que está sendo acessado
   Private lCn As IDbConnection     'Conexão que está sendo utilizada
   Private lTransaction As IDbTransaction 'Transação utilizada
   ' Private lTrans As IDbTransaction 'Objeto de transação 
#End Region

   Public ReadOnly Property Transaction() As IDbTransaction
      Get
         Return lTransaction
      End Get
   End Property

   Public ReadOnly Property Connection() As IDbConnection
      Get
         Return lCn
      End Get
   End Property

#Region "Construtores da classe"

   Public Sub New(ByVal XMLPath As String)
      ClConfig = New ClsConfig(XMLPath)
      StartPersiste()
   End Sub

   Public Sub New(ByVal oXML As Xml.XmlDocument)
      ClConfig = New ClsConfig(oXML)
      StartPersiste()
   End Sub

   Private Sub StartPersiste()
      Select Case ClConfig.DataBaseType
         Case ClsSQL.EnumDBType.SQLServer
            Dim Cn As New SqlClient.SqlConnection(ClConfig.ConnectionString)
            lCn = Cn
         Case ClsSQL.EnumDBType.OLEDB
            Dim Cn As New OleDb.OleDbConnection(ClConfig.ConnectionString)
            lCn = Cn
         Case ClsSQL.EnumDBType.ODBC
            Dim Cn As New Odbc.OdbcConnection(ClConfig.ConnectionString)
            lCn = Cn
         Case ClsSQL.EnumDBType.Oracle

         Case ClsSQL.EnumDBType.MySQL

      End Select
      If lCn.State = ConnectionState.Closed Then lCn.Open()
   End Sub

#End Region

#Region "Eventos"
   Public Event AfterInsert(ByVal Success As Boolean, ByVal Err As Exception)
   Public Event AfterUpdate(ByVal Success As Boolean, ByVal Err As Exception)
   Public Event AfterDelete(ByVal Success As Boolean, ByVal Err As Exception)
   Public Event AfterDeleteCriteria(ByVal Success As Boolean, ByVal RegAffect As Integer, ByVal Err As Exception)
   Public Event BeforeInsert(ByRef Cancel As Boolean)
   Public Event BeforeUpdate(ByRef Cancel As Boolean)
   Public Event BeforeDelete(ByRef Cancel As Boolean)
   Public Event BeforeDeleteCriteria(ByRef Cancel As Boolean)
#End Region

#Region "Métodos Privados"
   Private Function AbreConexao(ByVal iCn As IDbConnection) As IDbConnection
      If iCn.State = ConnectionState.Closed Then
         Dim XMLNodeList As Xml.XmlNodeList
         XMLNodeList = ReadXML("ConnectionString")
         iCn.ConnectionString = XMLNodeList.Item(0).InnerText
         iCn.Open()
      End If
      Return iCn
   End Function

   Private Function ReadXML(ByVal Table As String) As Xml.XmlNodeList
      Dim ObjXML As New Xml.XmlDocument
      Dim objNodeList As Xml.XmlNodeList
      ObjXML.Load(lXMLPath)
      objNodeList = ObjXML.GetElementsByTagName(Table)
      Return objNodeList
   End Function

   Private Function FieldsCharge(ByRef ObjValue As Object, ByVal objNodeList As Xml.XmlNodeList, ByVal lType As Type) As ClsSQL.stField()
      Dim i As Byte
      Dim Field() As ClsSQL.stField
      Dim TypeD As String
      Dim lPri As PropertyInfo
      ReDim Field(objNodeList.Item(0).ChildNodes.Count - 1)
      For i = 0 To objNodeList.Item(0).ChildNodes.Count - 1
         'Table Field Name         
         Field(i).TableFieldName = objNodeList.Item(0).ChildNodes(i).InnerText
         'Obj Field Name
         Field(i).ObjFieldName = objNodeList.Item(0).ChildNodes(i).Name
         'Primary Key Information
         Try
            Field(i).Key = (objNodeList.Item(0).ChildNodes(i).Attributes("Id").Value = "PK")
         Catch
            Field(i).Key = False
         End Try
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
         lPri = lType.GetProperty(Field(i).ObjFieldName)
         Field(i).Value = lPri.GetValue(ObjValue, Nothing)
      Next
      Return Field
   End Function

   Private Function MountOrder(ByVal Field() As ClsSQL.stField, ByVal Order As String) As String
      If Order = "" Then Return ""
      Dim i, j As Byte
      Dim ArrayOrder() As String
      Dim Ordert As New System.Text.StringBuilder
      ArrayOrder = Split(Order, ",")
      If ArrayOrder.Length > 0 Then
         For j = 0 To ArrayOrder.Length - 1
            For i = 0 To Field.Length - 1
               Dim Ord() As String
               Ord = Split(ArrayOrder(j).Trim, " ") 'Busca as palavras desc and asc no final da ordenação
               If Field(i).ObjFieldName.ToLower = Ord(0).ToLower Then
                  If Ordert.ToString <> "" Then Ordert.Append(",") 'Insere o separador para a ordenação
                  Ordert.Append(Field(i).TableFieldName)
                  If Ord.Length > 1 Then
                     Ordert.Append(" ")
                     Ordert.Append(Ord(1))
                  End If
                  Exit For
               End If
            Next
         Next
         Ordert.Insert(0, " order by ")
      End If
      Return (Ordert.ToString)
   End Function
#End Region

#Region "Métodos Públicos de Transação"
   Public Function BeginTransaction() As Boolean 'Inicia uma transação
      'Pensar melhor nos retornos
      If lStateTrans = EnumStateTrans.NotTransaction Then
         lTransaction = lCn.BeginTransaction
         lStateTrans = EnumStateTrans.InTransaction
         Return True
      Else
         Return False
      End If
   End Function

   Public Function Commit() As Boolean 'Confirma uma transação
      'Pensar melhor nos retornos
      If lStateTrans = EnumStateTrans.InTransaction Then
         lTransaction.Commit()
         lTransaction.Dispose()
         lStateTrans = EnumStateTrans.NotTransaction
         Return True
      Else
         Return False
      End If
   End Function

   Public Function Rollback() As Boolean 'Desfaz as operações realizadas pela transação
      'Pensar melhor nos retornos
      If lStateTrans = EnumStateTrans.InTransaction Then
         lTransaction.Rollback()
         lTransaction.Dispose()
         lStateTrans = EnumStateTrans.NotTransaction
         Return True
      Else
         Return False
      End If
   End Function
#End Region

#Region "Métodos Públicos de Procedures"

   Public Function ExecProcedureDirectValue(ByVal Procedure_Name As String, ByVal ParametersValues As String()) As Object
      Dim oProc As New Cls_Procedure(ClConfig, Procedure_Name, Cls_Procedure.e_Procedure_Type.ReturnDirectValue, lCn, lStateTrans, lTransaction, Cls_Procedure.e_Procedure_Type.ReturnDirectValue)
      Dim Parameter As PersisteDados.Cls_Procedure_Parameter
      Dim i As Integer = 0
      For Each Parameter In oProc.Collection_Parameters
         Parameter.Parameter_Value = ParametersValues(i)
         i += 1
      Next
      Return oProc.ExecuteSinglePropertyReader()
   End Function

   Public Function ExecProcedureNonQuery(ByVal Procedure_Name As String, ByVal ParametersValues As String()) As Integer
      Dim oProc As New Cls_Procedure(ClConfig, Procedure_Name, Cls_Procedure.e_Procedure_Type.NotReturnValue, lCn, lStateTrans, lTransaction, Cls_Procedure.e_Procedure_Type.NotReturnValue)
      Dim Parameter As PersisteDados.Cls_Procedure_Parameter
      Dim i As Integer = 0
      For Each Parameter In oProc.Collection_Parameters
         Parameter.Parameter_Value = ParametersValues(i)
         i += 1
      Next
      Return oProc.ExecuteNonQuery
   End Function

   Public Function ExecProcedureReader(ByVal Object_Proc As Object, ByVal ParametersValues As String()) As ArrayList
      Dim oProc As New Cls_Procedure(ClConfig, Object_Proc, lCn, lStateTrans, lTransaction, Cls_Procedure.e_Procedure_Type.ReturnValue)
      Dim Parameter As PersisteDados.Cls_Procedure_Parameter
      Dim i As Integer = 0
      For Each Parameter In oProc.Collection_Parameters
         Parameter.Parameter_Value = ParametersValues(i)
         i += 1
      Next
      Return oProc.ExecuteReader()
   End Function

#End Region


#Region "Métodos Públicos Insert/Update/Delete/DeleteCriteria"

   Public Sub Insert(ByVal objSave As Object)
      Dim Field() As ClsSQL.stField
      Dim SQLInsert As String
      Dim TableName As String
      Dim Prefix As String = ""

      TableName = ClConfig.DataBaseTableName(objSave, Prefix)  'Busca o nome da tabela no banco de dados
      Field = ClConfig.FieldsCharge(objSave)           'Monta a coleção de fields

      'Mounts the insert sql expression using ClsSQL class 
      Dim ClSQL As New ClsSQL(ClConfig.DataBaseType)
      SQLInsert = ClSQL.SQLInsert(TableName, Prefix, Field)
      ClSQL = Nothing

      Try
         'Execute expression and insert field in the database
         Dim Cancel As Boolean
         RaiseEvent BeforeInsert(Cancel)
         If Not Cancel Then
            Dim Cm As IDbCommand
            Select Case ClConfig.DataBaseType
               Case ClsSQL.EnumDBType.SQLServer
                  If lStateTrans = EnumStateTrans.InTransaction Then
                     Dim Cm1 As New SqlClient.SqlCommand(SQLInsert, lCn, DirectCast(lTransaction, SqlClient.SqlTransaction))
                     Cm = Cm1
                  Else
                     Dim Cm1 As New SqlClient.SqlCommand(SQLInsert, lCn)
                     Cm = Cm1
                  End If
               Case ClsSQL.EnumDBType.OLEDB
                  If lStateTrans = EnumStateTrans.InTransaction Then
                     Dim Cm1 As New OleDb.OleDbCommand(SQLInsert, lCn, DirectCast(lTransaction, OleDb.OleDbTransaction))
                     Cm = Cm1
                  Else
                     Dim Cm1 As New OleDb.OleDbCommand(SQLInsert, lCn)
                     Cm = Cm1
                  End If
               Case ClsSQL.EnumDBType.ODBC
                  If lStateTrans = EnumStateTrans.InTransaction Then
                     Dim Cm1 As New Odbc.OdbcCommand(SQLInsert, lCn, DirectCast(lTransaction, Odbc.OdbcTransaction))
                     Cm = Cm1
                  Else
                     Dim Cm1 As New Odbc.OdbcCommand(SQLInsert, lCn)
                     Cm = Cm1
                  End If
               Case ClsSQL.EnumDBType.Oracle
                  If lStateTrans = EnumStateTrans.InTransaction Then
                  Else
                  End If
               Case ClsSQL.EnumDBType.MySQL
                  If lStateTrans = EnumStateTrans.InTransaction Then
                  Else
                  End If
            End Select
            Cm.ExecuteNonQuery()
            Cm.Dispose()
            RaiseEvent AfterInsert(True, Nothing)
         End If
      Catch e As Exception
         RaiseEvent AfterInsert(False, e)
         Throw New Exception("Errors had occurred in Insert Operation.", e)
      End Try
   End Sub

   Public Sub Update(ByVal objSave As Object)
      Dim TableName As String
      Dim Field() As ClsSQL.stField
      Dim SQLUpdate As String, Prefix As String

      TableName = ClConfig.DataBaseTableName(objSave, Prefix)
      Field = ClConfig.FieldsCharge(objSave)

      'Mounts the insert sql expression using ClsSQL class 
      Dim ClSQL As New ClsSQL(lDataBase)
      SQLUpdate = ClSQL.SQLUpdate(TableName, Prefix, Field)
      ClSQL = Nothing

      'Executa a expressão de atualização, conforme tipo de banco de dados
      Try
         Dim Cancel As Boolean
         RaiseEvent BeforeUpdate(Cancel)
         If Not Cancel Then
            Dim Cm As IDbCommand
            Select Case ClConfig.DataBaseType
               Case ClsSQL.EnumDBType.SQLServer
                  If lStateTrans = EnumStateTrans.InTransaction Then
                     Dim Cm1 As New SqlClient.SqlCommand(SQLUpdate, lCn, DirectCast(lTransaction, SqlClient.SqlTransaction))
                     Cm = Cm1
                  Else
                     Dim Cm1 As New SqlClient.SqlCommand(SQLUpdate, lCn)
                     Cm = Cm1
                  End If
               Case ClsSQL.EnumDBType.OLEDB
                  If lStateTrans = EnumStateTrans.InTransaction Then
                     Dim Cm1 As New OleDb.OleDbCommand(SQLUpdate, lCn, DirectCast(lTransaction, OleDb.OleDbTransaction))
                     Cm = Cm1
                  Else
                     Dim Cm1 As New OleDb.OleDbCommand(SQLUpdate, lCn)
                     Cm = Cm1
                  End If
               Case ClsSQL.EnumDBType.ODBC
                  If lStateTrans = EnumStateTrans.InTransaction Then
                     Dim Cm1 As New Odbc.OdbcCommand(SQLUpdate, lCn, DirectCast(lTransaction, Odbc.OdbcTransaction))
                     Cm = Cm1
                  Else
                     Dim Cm1 As New Odbc.OdbcCommand(SQLUpdate, lCn)
                     Cm = Cm1
                  End If
               Case ClsSQL.EnumDBType.MySQL
                  If lStateTrans = EnumStateTrans.InTransaction Then
                  Else
                  End If
               Case ClsSQL.EnumDBType.Oracle
                  If lStateTrans = EnumStateTrans.InTransaction Then
                  Else
                  End If
            End Select
            Cm.ExecuteNonQuery()
            Cm.Dispose()
            RaiseEvent AfterUpdate(True, Nothing)
         End If
      Catch e As Exception
         RaiseEvent AfterUpdate(False, e)
         Throw New Exception("Errors had occurred in Update Operation.", e)
      End Try
   End Sub

   Public Sub Delete(ByVal objDelete As Object)
      Dim SQLDelete As String
      Dim TableName As String, Prefix As String
      Dim Field() As ClsSQL.stField

      TableName = ClConfig.DataBaseTableName(objDelete, Prefix)
      Field = ClConfig.FieldsCharge(objDelete)

      'Mounts the insert sql expression using ClsSQL class 
      Dim ClSQL As New ClsSQL(lDataBase)
      SQLDelete = ClSQL.SQLDelete(TableName, Prefix, Field)
      ClSQL = Nothing

      'Execute expression and delete field in the database
      Try
         Dim Cancel As Boolean
         RaiseEvent BeforeDelete(Cancel)
         If Not Cancel Then
            Dim Cm As IDbCommand
            Select Case ClConfig.DataBaseType
               Case ClsSQL.EnumDBType.SQLServer
                  If lStateTrans = EnumStateTrans.InTransaction Then
                     Dim Cm1 As New SqlClient.SqlCommand(SQLDelete, lCn, DirectCast(lTransaction, SqlClient.SqlTransaction))
                     Cm = Cm1
                  Else
                     Dim Cm1 As New SqlClient.SqlCommand(SQLDelete, lCn)
                     Cm = Cm1
                  End If
               Case ClsSQL.EnumDBType.OLEDB
                  If lStateTrans = EnumStateTrans.InTransaction Then
                     Dim Cm1 As New OleDb.OleDbCommand(SQLDelete, lCn, DirectCast(lTransaction, OleDb.OleDbTransaction))
                     Cm = Cm1
                  Else
                     Dim Cm1 As New OleDb.OleDbCommand(SQLDelete, lCn)
                     Cm = Cm1
                  End If
               Case ClsSQL.EnumDBType.ODBC
                  If lStateTrans = EnumStateTrans.InTransaction Then
                     Dim Cm1 As New Odbc.OdbcCommand(SQLDelete, lCn, DirectCast(lTransaction, Odbc.OdbcTransaction))
                     Cm = Cm1
                  Else
                     Dim Cm1 As New Odbc.OdbcCommand(SQLDelete, lCn)
                     Cm = Cm1
                  End If
               Case ClsSQL.EnumDBType.Oracle
                  If lStateTrans = EnumStateTrans.InTransaction Then
                  Else
                  End If
               Case ClsSQL.EnumDBType.MySQL
                  If lStateTrans = EnumStateTrans.InTransaction Then
                  Else
                  End If
            End Select
            Cm.ExecuteNonQuery()
            Cm.Dispose()
            RaiseEvent AfterDelete(True, Nothing)
         End If
      Catch e As Exception
         RaiseEvent AfterDelete(False, e)
         Throw New Exception("Errors had occurred in Delete Operation.", e)
      End Try
   End Sub

   Public Function DeleteCriteria(ByVal objDelete As Object, ByVal Criteria As ClsCriteria) As Integer
      Dim SQLDelete As String
      Dim TableName As String, Prefix As String
      Dim Field() As ClsSQL.stField
      Dim RegAffect As Integer

      TableName = ClConfig.DataBaseTableName(objDelete, Prefix)
      Field = ClConfig.FieldsCharge(objDelete)

      'Mounts the insert sql expression using ClsSQL class 
      Dim ClSQL As New ClsSQL(lDataBase)
      SQLDelete = ClSQL.SQLDeleteCriteria(TableName, Prefix, Criteria)
      ClSQL = Nothing

      'Execute expression and delete field in the database
      Try
         Dim Cancel As Boolean
         RaiseEvent BeforeDeleteCriteria(Cancel)
         If Not Cancel Then
            Dim Cm As IDbCommand
            Select Case ClConfig.DataBaseType
               Case ClsSQL.EnumDBType.SQLServer
                  If lStateTrans = EnumStateTrans.InTransaction Then
                     Dim Cm1 As New SqlClient.SqlCommand(SQLDelete, lCn, DirectCast(lTransaction, SqlClient.SqlTransaction))
                     Cm = Cm1
                  Else
                     Dim Cm1 As New SqlClient.SqlCommand(SQLDelete, lCn)
                     Cm = Cm1
                  End If
               Case ClsSQL.EnumDBType.OLEDB
                  If lStateTrans = EnumStateTrans.InTransaction Then
                     Dim Cm1 As New OleDb.OleDbCommand(SQLDelete, lCn, DirectCast(lTransaction, OleDb.OleDbTransaction))
                     Cm = Cm1
                  Else
                     Dim Cm1 As New OleDb.OleDbCommand(SQLDelete, lCn)
                     Cm = Cm1
                  End If
               Case ClsSQL.EnumDBType.ODBC
                  If lStateTrans = EnumStateTrans.InTransaction Then
                     Dim Cm1 As New Odbc.OdbcCommand(SQLDelete, lCn, DirectCast(lTransaction, Odbc.OdbcTransaction))
                     Cm = Cm1
                  Else
                     Dim Cm1 As New Odbc.OdbcCommand(SQLDelete, lCn)
                     Cm = Cm1
                  End If
               Case ClsSQL.EnumDBType.Oracle
                  If lStateTrans = EnumStateTrans.InTransaction Then
                  Else
                  End If
               Case ClsSQL.EnumDBType.MySQL
                  If lStateTrans = EnumStateTrans.InTransaction Then
                  Else
                  End If
            End Select
            RegAffect = Cm.ExecuteNonQuery()
            Cm.Dispose()
            RaiseEvent AfterDeleteCriteria(True, RegAffect, Nothing)
         End If
         Return RegAffect
      Catch e As Exception
         RaiseEvent AfterDeleteCriteria(False, RegAffect, e)
         Throw New Exception("Errors had occurred in Delete Operation.", e)
      End Try
   End Function

#End Region

#Region "Métodos Públicos de Seleção"

   Public Function ObjectReturnPk(ByVal ObjReturn As Object) As Object
      'Retorna um registro a partir dos valores passados na chave primária
      'ObjReturn = Objeto que será retornado. TEm que vir com as propriedades correspondentes a chave primária preenchidas

      Dim TableName As String, Prefix As String = ""
      Dim Field() As ClsSQL.stField
      Dim SQLWhere As String
      Dim i As Byte = 0
      Dim lType As Type
      Dim lPri As PropertyInfo
      Dim Dr As IDataReader

      lType = ObjReturn.GetType

      'Get the database table name
      TableName = ClConfig.DataBaseTableName(ObjReturn, Prefix)

      'Mounts the field array object
      Field = ClConfig.FieldsCharge(ObjReturn) 'FieldsCharge(ObjReturn, objNodeList, lType)

      'Mont the SQL expression for fields select
      Dim ClSQL As New ClsSQL(lDataBase)
      SQLWhere = ClSQL.SQLWhere(Field)
      ClSQL = Nothing

      'Mont the complete SQL selection expression 
      Dim Sel As New System.Text.StringBuilder("Select * from ")
      Sel.Append(Prefix)
      Sel.Append(TableName)
      Sel.Append(" as ")
      Sel.Append(TableName)
      Sel.Append(" ")
      Sel.Append(SQLWhere)

      'Select the fields from DataBase
      Select Case ClConfig.DataBaseType
         Case ClsSQL.EnumDBType.SQLServer
            Dim Cm As SqlClient.SqlCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New SqlClient.SqlCommand(Sel.ToString, lCn, lTransaction)
            Else
               Cm = New SqlClient.SqlCommand(Sel.ToString, lCn)
            End If

            Dim lDr As SqlClient.SqlDataReader
            lDr = Cm.ExecuteReader
            Dr = lDr
            Cm.Dispose()
         Case ClsSQL.EnumDBType.OLEDB
            Dim Cm As OleDb.OleDbCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New OleDb.OleDbCommand(Sel.ToString, lCn, lTransaction)
            Else
               Cm = New OleDb.OleDbCommand(Sel.ToString, lCn)
            End If
            Dim lDr As OleDb.OleDbDataReader
            lDr = Cm.ExecuteReader
            Dr = lDr
            Cm.Dispose()
         Case ClsSQL.EnumDBType.ODBC
            Dim Cm As Odbc.OdbcCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New Odbc.OdbcCommand(Sel.ToString, lCn, lTransaction)
            Else
               Cm = New Odbc.OdbcCommand(Sel.ToString, lCn)
            End If
            Dim lDr As Odbc.OdbcDataReader
            lDr = Cm.ExecuteReader
            Dr = lDr
            Cm.Dispose()
         Case ClsSQL.EnumDBType.Oracle

         Case ClsSQL.EnumDBType.MySQL

      End Select
      Sel = Nothing
      If Dr.Read Then
         For i = 0 To Field.Length - 1
            lPri = lType.GetProperty(Field(i).ObjFieldName)
            If Dr.Item(Field(i).TableFieldName).Equals(System.DBNull.Value) Then
               lPri.SetValue(ObjReturn, Nothing, Nothing)
            Else
               lPri.SetValue(ObjReturn, Dr.Item(Field(i).TableFieldName), Nothing)
            End If
         Next
         Dr.Close()
         Return ObjReturn
      Else
         Dr.Close()
         Return Nothing
      End If
   End Function

   Public Function ObjectsReturn(ByVal ObjWhere As Object, Optional ByVal Order As String = "") As ArrayList
      'Retorna uma coleção de objetos, a partir do objeto passado no objWhere
      'O objeto passado será utilizado para montagem dos critérios de seleção, com as propriedades que estiverem preenchidas
      'ObjWhere = Objeto de critério e que será retornado dentro do arraylist
      'Order = Propriedade do objeto que indexará a seleção
      Dim TableName As String, Prefix As String = ""
      Dim Field() As ClsSQL.stField

      Dim i As Byte = 0
      Dim DrItens As IDataReader
      Dim Ass As System.Reflection.Assembly
      Dim lType As Type
      Dim lPri As PropertyInfo
      Dim lReturn As New ArrayList

      lType = ObjWhere.GetType

      TableName = ClConfig.DataBaseTableName(ObjWhere, Prefix)
      Field = ClConfig.FieldsCharge(ObjWhere)

      Dim SQLWhere As New System.Text.StringBuilder("Select * from ")
      SQLWhere.Append(Prefix)
      SQLWhere.Append(TableName)
      SQLWhere.Append(" ")

      'Monts the SQL expression for fields select
      Dim ClSQL As New ClsSQL(lDataBase)
      SQLWhere.Append(ClSQL.SQLWhere(Field, True, False))
      ClSQL = Nothing

      'Monta a cláusula do order by
      SQLWhere.Append(MountOrder(Field, Order))

      Select Case ClConfig.DataBaseType
         Case ClsSQL.EnumDBType.SQLServer
            Dim Cm As SqlClient.SqlCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New SqlClient.SqlCommand(SQLWhere.ToString, lCn, lTransaction)
            Else
               Cm = New SqlClient.SqlCommand(SQLWhere.ToString, lCn)
            End If
            Dim Dr As SqlClient.SqlDataReader
            Dr = Cm.ExecuteReader
            DrItens = Dr
            Cm.Dispose()
         Case ClsSQL.EnumDBType.OLEDB
            Dim Cm As OleDb.OleDbCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New OleDb.OleDbCommand(SQLWhere.ToString, lCn, lTransaction)
            Else
               Cm = New OleDb.OleDbCommand(SQLWhere.ToString, lCn)
            End If
            Dim Dr As OleDb.OleDbDataReader
            Dr = Cm.ExecuteReader
            DrItens = Dr
            Cm.Dispose()
         Case ClsSQL.EnumDBType.ODBC
            Dim Cm As Odbc.OdbcCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New Odbc.OdbcCommand(SQLWhere.ToString, lCn, lTransaction)
            Else
               Cm = New Odbc.OdbcCommand(SQLWhere.ToString, lCn)
            End If

            Dim Dr As Odbc.OdbcDataReader
            Dr = Cm.ExecuteReader
            DrItens = Dr
            Cm.Dispose()
         Case ClsSQL.EnumDBType.Oracle

         Case ClsSQL.EnumDBType.MySQL

      End Select

      Ass = lType.Assembly
      Do While DrItens.Read()
         Dim Obj As New Object
         Obj = Ass.CreateInstance(lType.FullName)
         For i = 0 To Field.Length - 1
            lPri = lType.GetProperty(Field(i).ObjFieldName)
            If DrItens.Item(Field(i).TableFieldName).Equals(System.DBNull.Value) Then
               lPri.SetValue(Obj, Nothing, Nothing)
            Else
               lPri.SetValue(Obj, DrItens.Item(Field(i).TableFieldName), Nothing)
            End If
         Next
         lReturn.Add(Obj)
      Loop
      DrItens.Close()
      DrItens.Dispose()
      Return lReturn
   End Function

   Public Function AllObjects(ByVal ObjReturn As Object, Optional ByVal Order As String = "") As ArrayList
      'Retorna um ArrayList com todos os registros da tabela correspondente ao objeto
      'ObjReturn = Objeto que será retornado dentro do arraylist
      'Order     = Parametro de ordenação. Vem com o nome da propriedade do objeto e não com o nome do campo na tabela

      Dim TableName As String, Prefix As String = ""
      Dim i As Byte
      Dim lType As Type 'Tipo do objeto passado
      Dim lPri As PropertyInfo
      Dim DrItens As IDataReader
      Dim Field() As ClsSQL.stField

      Dim Ass As System.Reflection.Assembly
      Dim lReturn As New ArrayList

      lType = ObjReturn.GetType

      'Busca o nome da tabela no banco de dados
      TableName = ClConfig.DataBaseTableName(ObjReturn, Prefix)

      'Monta a expressão SQL de seleção
      Dim SQL As New System.Text.StringBuilder("Select * from ")
      SQL.Append(Prefix)
      SQL.Append(TableName)
      SQL.Append(" as ")
      SQL.Append(TableName)
      Field = ClConfig.FieldsCharge(ObjReturn)

      'Monta a cláusula do order by
      SQL.Append(MountOrder(Field, Order))

      Select Case ClConfig.DataBaseType 'Executa o comando de leitura, de acordo com o tipo de banco de dados
         Case ClsSQL.EnumDBType.SQLServer
            Dim Dr As SqlClient.SqlDataReader
            Dim Cm As SqlClient.SqlCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New SqlClient.SqlCommand(SQL.ToString, lCn, lTransaction)
            Else
               Cm = New SqlClient.SqlCommand(SQL.ToString, lCn)
            End If
            Dr = Cm.ExecuteReader()
            Cm.Dispose()
            DrItens = Dr
         Case ClsSQL.EnumDBType.OLEDB
            Dim Dr As OleDb.OleDbDataReader
            Dim Cm As OleDb.OleDbCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New OleDb.OleDbCommand(SQL.ToString, lCn, lTransaction)
            Else
               Cm = New OleDb.OleDbCommand(SQL.ToString, lCn)
            End If
            Dr = Cm.ExecuteReader()
            Cm.Dispose()
            DrItens = Dr
         Case ClsSQL.EnumDBType.ODBC
            Dim Dr As Odbc.OdbcDataReader
            Dim Cm As Odbc.OdbcCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New Odbc.OdbcCommand(SQL.ToString, lCn, lTransaction)
            Else
               Cm = New Odbc.OdbcCommand(SQL.ToString, lCn)
            End If
            Dr = Cm.ExecuteReader()
            Cm.Dispose()
            DrItens = Dr
         Case ClsSQL.EnumDBType.Oracle

         Case ClsSQL.EnumDBType.MySQL

      End Select

      Ass = lType.Assembly
      Do While DrItens.Read()
         'Cria uma instancia do objeto recebido
         Dim Obj As New Object
         Obj = Ass.CreateInstance(lType.FullName)

         For i = 0 To Field.Length - 1
            lPri = lType.GetProperty(Field(i).ObjFieldName)

            If DrItens.Item(Field(i).TableFieldName).Equals(System.DBNull.Value) Then
               lPri.SetValue(Obj, Nothing, Nothing)
            Else
               lPri.SetValue(Obj, DrItens.Item(Field(i).TableFieldName), Nothing)
            End If
         Next
         lReturn.Add(Obj)

      Loop
      DrItens.Close()
      DrItens.Dispose()
      Return lReturn
   End Function

   Public Function SingleObjects(ByVal objReturn As Object, ByVal objCriteria As ClsCriteria, Optional ByVal Order As String = "") As ArrayList
      'Rotina que retorna uma coleção de objetos 
      'Retorna uma coleção de objetos, a partir do objeto passado no objWhere
      'O objeto passado será utilizado para montagem dos critérios de seleção, com as propriedades que estiverem preenchidas
      'ObjWhere = Objeto de critério e que será retornado dentro do arraylist
      'Order = Propriedade do objeto que indexará a seleção (Pode-se passar várias propriedades, separados por vírgula
      Dim TableName As String, Prefix As String = ""
      Dim lType As Type 'Tipo do objeto passado
      Dim lPri As PropertyInfo
      Dim DrItens As IDataReader
      Dim Field() As ClsSQL.stField
      Dim lReturn As New ArrayList
      Dim i As Byte = 0
      Dim Ass As System.Reflection.Assembly
      lType = objReturn.GetType
      TableName = ClConfig.DataBaseTableName(lType.Name.ToString, Prefix)
      Field = ClConfig.FieldsCharge(objReturn)

      Dim SQLWhere As New System.Text.StringBuilder("Select * from ")
      SQLWhere.Append(Prefix)
      SQLWhere.Append(TableName)
      SQLWhere.Append(" as ")
      SQLWhere.Append(TableName)
      SQLWhere.Append(objCriteria.SQLWhere) 'Busca a cláusula where do objeto

      'Monta a cláusula do order by
      SQLWhere.Append(MountOrder(Field, Order))

      Select Case ClConfig.DataBaseType
         Case ClsSQL.EnumDBType.SQLServer
            Dim Cm As SqlClient.SqlCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New SqlClient.SqlCommand(SQLWhere.ToString, lCn, lTransaction)
            Else
               Cm = New SqlClient.SqlCommand(SQLWhere.ToString, lCn)
            End If
            Dim Dr As SqlClient.SqlDataReader
            Dr = Cm.ExecuteReader
            DrItens = Dr
            Cm.Dispose()
         Case ClsSQL.EnumDBType.OLEDB
            Dim Cm As OleDb.OleDbCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New OleDb.OleDbCommand(SQLWhere.ToString, lCn, lTransaction)
            Else
               Cm = New OleDb.OleDbCommand(SQLWhere.ToString, lCn)
            End If
            Dim Dr As OleDb.OleDbDataReader
            Dr = Cm.ExecuteReader
            DrItens = Dr
            Cm.Dispose()
         Case ClsSQL.EnumDBType.ODBC
            Dim Cm As Odbc.OdbcCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New Odbc.OdbcCommand(SQLWhere.ToString, lCn, lTransaction)
            Else
               Cm = New Odbc.OdbcCommand(SQLWhere.ToString, lCn)
            End If
            Dim Dr As Odbc.OdbcDataReader
            Dr = Cm.ExecuteReader
            DrItens = Dr
            Cm.Dispose()
         Case ClsSQL.EnumDBType.Oracle

         Case ClsSQL.EnumDBType.MySQL

      End Select

      Ass = lType.Assembly
      Do While DrItens.Read()
         Dim Obj As New Object
         Obj = Ass.CreateInstance(lType.FullName)
         For i = 0 To Field.Length - 1
            lPri = lType.GetProperty(Field(i).ObjFieldName)
            If DrItens.Item(Field(i).TableFieldName).ToString.Trim = "" Then
               lPri.SetValue(Obj, Nothing, Nothing)
            Else
               lPri.SetValue(Obj, DrItens.Item(Field(i).TableFieldName), Nothing)
            End If
         Next
         lReturn.Add(Obj)
      Loop
      DrItens.Close()
      DrItens.Dispose()
      Return lReturn
   End Function

#End Region

#Region "Métodos Públicos Auxiliares (NextSingleKeyNumber/ObjectSingleCount)"

   Public Function NextSingleKeyNumber(ByVal Obj As Object, ByVal PropertyName As String, Optional ByVal Criteria As ClsCriteria = Nothing) As Long
      Dim lType As Type 'Tipo do objeto passado
      Dim Prefix As String = "", TableName As String, FieldName As String
      Dim SQL As String, DrItens As IDataReader
      lType = Obj.GetType
      TableName = ClConfig.DataBaseTableName(lType.Name.ToString, Prefix)
      FieldName = ClConfig.GetFieldName(lType.Name.ToString, PropertyName)
      Dim ClSQL As New ClsSQL(lDataBase)
      SQL = ClSQL.MontSQLTop(TableName, FieldName, Criteria)

      Select Case ClConfig.DataBaseType 'Executa o comando de leitura, de acordo com o tipo de banco de dados
         Case ClsSQL.EnumDBType.SQLServer
            Dim Dr As SqlClient.SqlDataReader
            Dim Cm As SqlClient.SqlCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New SqlClient.SqlCommand(SQL, lCn, lTransaction)
            Else
               Cm = New SqlClient.SqlCommand(SQL, lCn)
            End If
            Dr = Cm.ExecuteReader()
            Cm.Dispose()
            DrItens = Dr
         Case ClsSQL.EnumDBType.OLEDB
            Dim Dr As OleDb.OleDbDataReader
            Dim Cm As OleDb.OleDbCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New OleDb.OleDbCommand(SQL, lCn, lTransaction)
            Else
               Cm = New OleDb.OleDbCommand(SQL, lCn)
            End If

            Dr = Cm.ExecuteReader()
            Cm.Dispose()
            DrItens = Dr
         Case ClsSQL.EnumDBType.ODBC
            Dim Dr As Odbc.OdbcDataReader
            Dim Cm As Odbc.OdbcCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New Odbc.OdbcCommand(SQL, lCn, lTransaction)
            Else
               Cm = New Odbc.OdbcCommand(SQL, lCn)
            End If
            Dr = Cm.ExecuteReader()
            Cm.Dispose()
            DrItens = Dr
         Case ClsSQL.EnumDBType.Oracle

         Case ClsSQL.EnumDBType.MySQL

      End Select
      Dim Response As Long
      If DrItens.Read Then
         If DrItens.Item(0).Equals(System.DBNull.Value) Then
            Response = 1
         Else
            Response = Long.Parse(DrItens(0)) + 1
         End If

      Else
         Response = 1
      End If
      DrItens.Dispose()
      DrItens = Nothing
      Return Response
   End Function

   Public Function ObjectSingleCount(ByVal Obj As Object, ByVal PropertyName As String, Optional ByVal Criteria As ClsCriteria = Nothing) As Long
      'Retorna um count sem agrupamento por uma propriedade qualquer do objeto
      Dim lType As Type
      Dim Prefix As String = "", TableName As String, FieldName As String
      Dim SQL As String, DrItens As IDataReader
      lType = Obj.GetType
      TableName = ClConfig.DataBaseTableName(lType.Name.ToString, Prefix)
      FieldName = ClConfig.GetFieldName(lType.Name.ToString, PropertyName)
      Dim ClSQL As New ClsSQL(lDataBase)
      SQL = ClSQL.MontSQLSingleCount(TableName, FieldName, Criteria)

      Select Case ClConfig.DataBaseType 'Executa o comando de leitura, de acordo com o tipo de banco de dados
         Case ClsSQL.EnumDBType.SQLServer
            Dim Dr As SqlClient.SqlDataReader
            Dim Cm As SqlClient.SqlCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New SqlClient.SqlCommand(SQL, lCn, lTransaction)
            Else
               Cm = New SqlClient.SqlCommand(SQL, lCn)
            End If
            Dr = Cm.ExecuteReader()
            Cm.Dispose()
            DrItens = Dr
         Case ClsSQL.EnumDBType.OLEDB
            Dim Dr As OleDb.OleDbDataReader
            Dim Cm As OleDb.OleDbCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New OleDb.OleDbCommand(SQL, lCn, lTransaction)
            Else
               Cm = New OleDb.OleDbCommand(SQL, lCn)
            End If

            Dr = Cm.ExecuteReader()
            Cm.Dispose()
            DrItens = Dr
         Case ClsSQL.EnumDBType.ODBC
            Dim Dr As Odbc.OdbcDataReader
            Dim Cm As Odbc.OdbcCommand
            If lStateTrans = EnumStateTrans.InTransaction Then
               Cm = New Odbc.OdbcCommand(SQL, lCn, lTransaction)
            Else
               Cm = New Odbc.OdbcCommand(SQL, lCn)
            End If
            Dr = Cm.ExecuteReader()
            Cm.Dispose()
            DrItens = Dr
         Case ClsSQL.EnumDBType.Oracle

         Case ClsSQL.EnumDBType.MySQL

      End Select
      Dim Response As Long
      If DrItens.Read Then
         If DrItens.Item(0).Equals(System.DBNull.Value) Then
            Response = 0
         Else
            Response = Long.Parse(DrItens(0))
         End If

      Else
         Response = 1
      End If
      DrItens.Dispose()
      DrItens = Nothing
      Return Response
   End Function

#End Region

#Region "Finalizadores"

   Private Sub LiberaObjetos()
      Try
         'If Not IsNothing(lTrans) Then
         ' lTrans.Dispose()
         ' lTrans = Nothing
         ' End If
         If Not IsNothing(lTransaction) Then
            lTransaction.Dispose()
            lTransaction = Nothing
         End If
         If Not IsNothing(lCn) Then
            If lCn.State = ConnectionState.Open Then
               lCn.Close()
            End If
            lCn = Nothing
         End If
         ClConfig = Nothing
      Catch e As Exception
      End Try
      ClConfig = Nothing
   End Sub

   Protected Overrides Sub Finalize()
      LiberaObjetos()
      MyBase.Finalize()
   End Sub

   Public Sub Dispose() Implements System.IDisposable.Dispose
      LiberaObjetos()
      System.GC.SuppressFinalize(Me)
   End Sub
#End Region


End Class
