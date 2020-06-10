Imports System.Reflection

Friend Class Cls_Procedure

   Public Enum EnumStateTrans
      NotTransaction = 0
      InTransaction = 1
   End Enum

   Public Enum e_Procedure_Type
      NotReturnValue = 0
      ReturnValue = 1
      ReturnDirectValue = 2
   End Enum

   Private lStateTrans As EnumStateTrans = EnumStateTrans.NotTransaction
   Private lTrans As IDbTransaction

   Private Cn As IDbConnection

   Private str_Procedure_Name As String 'NOME DA PROCEDURE NO BANCO DE DADOS
   Private obj_Object_Name As String    'NOME DO OBJETO PASSADO
   Private ObjProcedure As Object       'OBJETO QUE RECEBERÁ OS VALORES RETORNADOS PELA PROCEDURE

   Private obj_Collection_Fields As New Cls_Collection_Procedure_Field         'COLEÇÃO DE CAMPOS DA PROCEDURE
   Private obj_Collection_Parameters As New Cls_Collection_Procedure_Parameter 'COLEÇÃO DE PARÂMETROS DA PROCEDURE

   Private int_RegAffect As Integer                    'QUANTIDADE DE REGISTROS AFETADOS PELA PROCEDURE
   Private obj_Procedure_Return_Direct_Value As Object 'VALOR RETORNADO PELA PROCEDURE (PROCEDURE DE UM ÚNICO CAMPO, COM UM ÚNICO REGISTRO
   Private obj_Procedure_Return_Rows As ArrayList      'COLEÇÃO DE OBJETOS RETORNADOS PELA PROCEDURE

   Private objConfig As ClsConfig                      'CLASSE DE CONFIGURAÇÃO PASSADO PARA A PROCEDURE

   Private noProcedure As Xml.XmlNode

   Private int_Procedure_Type As e_Procedure_Type 'TIPO DA PROCEDURE QUE ESTÁ SENDO EXECUTADA

   Public Sub New(ByVal oConfig As ClsConfig, ByVal Obj_Procedure As Object, ByVal iCn As IDbConnection, ByVal iStateTrans As EnumStateTrans, ByVal iTrans As IDbTransaction, ByVal lProcedure_type As e_Procedure_Type)
      'Sempre retorna coleção, porque o retorno será dado no objeto passado
      objConfig = oConfig
      ObjProcedure = Obj_Procedure
      int_Procedure_Type = lProcedure_type
      Cn = iCn
      lStateTrans = iStateTrans
      lTrans = iTrans
      str_Procedure_Name = Get_Procedure_Details(Obj_Procedure)
      
   End Sub

   Public Sub New(ByVal oConfig As ClsConfig, ByVal str_Procedure As String, ByVal Procedure_Type As e_Procedure_Type, ByVal iCn As IDbConnection, ByVal iStateTrans As EnumStateTrans, ByVal iTrans As IDbTransaction, ByVal lProcedure_type As e_Procedure_Type)
      'Sempre retorna coleção, porque o retorno será dado no objeto passado
      objConfig = oConfig
      int_Procedure_Type = lProcedure_type
      Cn = iCn
      lStateTrans = iStateTrans
      lTrans = iTrans
      str_Procedure_Name = Get_Procedure_Details(str_Procedure)
   End Sub

   Private Function Get_Procedure_Details(ByVal str_Procedure As String) As String
      Dim proc_name As String = ""
      noProcedure = objConfig.GetProcedureDetails(str_Procedure)
      If noProcedure.ChildNodes.Count = 0 Then
         Throw New Exception("Não foi possível resolver o nome da procedure.")
      Else
         proc_name = noProcedure.Attributes("procedure_name").Value
      End If
      Mont_Details()
      Return proc_name
   End Function

   Private Function Get_Procedure_Details(ByVal Obj As Object) As String
      Dim lType As Type = Obj.GetType
      Dim proc_name As String = Get_Procedure_Details(lType.Name.ToString())
      Return proc_name
   End Function

   Private Sub Mont_Details()
      Dim no As Xml.XmlNode
      Dim noParameter As Byte = 1
      If int_Procedure_Type <> e_Procedure_Type.NotReturnValue Then
         For Each no In noProcedure.ChildNodes(0).ChildNodes
            'PROPRIEDADES
            Dim oProp As New Cls_Procedure_Field
            oProp.Field_Name = no.Attributes("FieldName").Value
            oProp.Property_Name = no.Attributes("PropertyName").Value
            obj_Collection_Fields.Add(oProp)
         Next
      Else
         noParameter = 0
      End If
      For Each no In noProcedure.ChildNodes(noParameter).ChildNodes
         'Parametros
         Dim oPar As New Cls_Procedure_Parameter
         oPar.Parameter_Name = no.Attributes("ParDatabaseName").Value
         'oPar.Property_Name = no.Attributes("ParObjectName").Value
         Select Case no.Attributes("Par_Type").Value.ToLower
            Case "text"
               oPar.Parameter_Type = Cls_Procedure_Parameter.EnumType.eText
            Case "integer"
               oPar.Parameter_Type = Cls_Procedure_Parameter.EnumType.eInteger
            Case "float"
               oPar.Parameter_Type = Cls_Procedure_Parameter.EnumType.eFloat
            Case "date"
               oPar.Parameter_Type = Cls_Procedure_Parameter.EnumType.eDate
            Case "time"
               oPar.Parameter_Type = Cls_Procedure_Parameter.EnumType.eTime
            Case "datetime"
               oPar.Parameter_Type = Cls_Procedure_Parameter.EnumType.eDateTime
            Case "bool"
               oPar.Parameter_Type = Cls_Procedure_Parameter.EnumType.eBool
         End Select
         obj_Collection_Parameters.Add(oPar)
      Next
   End Sub

   Public Function ExecuteNonQuery() As Integer
      'Para procedures que só retornam a quantidade de registros que será afetado
      If Procedure_Type <> e_Procedure_Type.NotReturnValue Then
         Throw New Exception("The procedure Type don't is configuration correctly!") 'Montar a mensagem em inglês
      End If

      Dim SQL As String = Mont_SQL_Exec_Procedure()
      Dim Cm As IDbCommand
      If lStateTrans = EnumStateTrans.NotTransaction Then
         Select Case objConfig.DataBaseType
            Case ClsSQL.EnumDBType.SQLServer
               Dim Cm1 As New SqlClient.SqlCommand(SQL, Cn)
               Cm = Cm1
            Case ClsSQL.EnumDBType.OLEDB
               Dim Cm1 As New OleDb.OleDbCommand(SQL, Cn)
               Cm = Cm1
            Case ClsSQL.EnumDBType.ODBC
               Dim Cm1 As New Odbc.OdbcCommand(SQL, Cn)
               Cm = Cm1
            Case ClsSQL.EnumDBType.Oracle
            Case ClsSQL.EnumDBType.MySQL
         End Select
      Else
         Select Case objConfig.DataBaseType
            Case ClsSQL.EnumDBType.SQLServer
               Dim Cm1 As New SqlClient.SqlCommand(SQL, Cn, lTrans)
               Cm = Cm1
            Case ClsSQL.EnumDBType.OLEDB
               Dim Cm1 As New OleDb.OleDbCommand(SQL, Cn, lTrans)
               Cm = Cm1
            Case ClsSQL.EnumDBType.ODBC
               Dim Cm1 As New Odbc.OdbcCommand(SQL, Cn, lTrans)
               Cm = Cm1
            Case ClsSQL.EnumDBType.Oracle
            Case ClsSQL.EnumDBType.MySQL
         End Select
      End If
      Return Cm.ExecuteNonQuery()
   End Function

   Public Function ExecuteReader() As ArrayList
      'Para procedures que retornam coleção de registros

      If int_Procedure_Type <> e_Procedure_Type.ReturnValue Then
         Throw New Exception("The procedure Type don't is configuration correctly!")
      End If

      Dim SQL As String = Mont_SQL_Exec_Procedure()
      Dim Cm As IDbCommand
      Dim DrItens As IDataReader
      If lStateTrans = EnumStateTrans.NotTransaction Then
         Select Case objConfig.DataBaseType
            Case ClsSQL.EnumDBType.SQLServer
               Dim Cm1 As New SqlClient.SqlCommand(SQL, Cn)
               Cm = Cm1
            Case ClsSQL.EnumDBType.OLEDB
               Dim Cm1 As New OleDb.OleDbCommand(SQL, Cn)
               Cm = Cm1
            Case ClsSQL.EnumDBType.ODBC
               Dim Cm1 As New Odbc.OdbcCommand(SQL, Cn)
               Cm = Cm1
            Case ClsSQL.EnumDBType.Oracle
            Case ClsSQL.EnumDBType.MySQL
         End Select
      Else
         Select Case objConfig.DataBaseType
            Case ClsSQL.EnumDBType.SQLServer
               Dim Cm1 As New SqlClient.SqlCommand(SQL, Cn, lTrans)
               Cm = Cm1
            Case ClsSQL.EnumDBType.OLEDB
               Dim Cm1 As New OleDb.OleDbCommand(SQL, Cn, lTrans)
               Cm = Cm1
            Case ClsSQL.EnumDBType.ODBC
               Dim Cm1 As New Odbc.OdbcCommand(SQL, Cn, lTrans)
               Cm = Cm1
            Case ClsSQL.EnumDBType.Oracle
            Case ClsSQL.EnumDBType.MySQL
         End Select
      End If
      DrItens = Cm.ExecuteReader()
      Dim lType As Type
      lType = ObjProcedure.GetType
      Dim Ass As Assembly = lType.Assembly
      Dim lPri As PropertyInfo
      Dim lReturn As New ArrayList
      Dim objField As Cls_Procedure_Field
      Do While DrItens.Read()
         'Cria uma instancia do objeto recebido
         Dim Obj As New Object
         Obj = Ass.CreateInstance(lType.FullName)

         For Each objField In obj_Collection_Fields
            lPri = lType.GetProperty(objField.Property_Name)
            If DrItens.Item(objField.Field_Name).Equals(System.DBNull.Value) Then
               lPri.SetValue(Obj, Nothing, Nothing)
            Else
               lPri.SetValue(Obj, DrItens.Item(objField.Field_Name), Nothing)
            End If
         Next
         lReturn.Add(Obj)
      Loop
      DrItens.Close()
      DrItens.Dispose()
      Return lReturn
   End Function

   Public Function ExecuteSinglePropertyReader() As Object
      'Para procedures que retornam um único campo e um único registro como resultado
      If Procedure_Type <> e_Procedure_Type.ReturnDirectValue Then
         Throw New Exception("The procedure Type don't is configuration correctly!")
      End If

      Dim SQL As String = Mont_SQL_Exec_Procedure()
      Dim Cm As IDbCommand
      Dim DrItens As IDataReader

      If lStateTrans = EnumStateTrans.NotTransaction Then
         Select Case objConfig.DataBaseType
            Case ClsSQL.EnumDBType.SQLServer
               Dim Cm1 As New SqlClient.SqlCommand(SQL, Cn)
               Cm = Cm1
            Case ClsSQL.EnumDBType.OLEDB
               Dim Cm1 As New OleDb.OleDbCommand(SQL, Cn)
               Cm = Cm1
            Case ClsSQL.EnumDBType.ODBC
               Dim Cm1 As New Odbc.OdbcCommand(SQL, Cn)
               Cm = Cm1
            Case ClsSQL.EnumDBType.Oracle
            Case ClsSQL.EnumDBType.MySQL
         End Select
      Else
         Select Case objConfig.DataBaseType
            Case ClsSQL.EnumDBType.SQLServer
               Dim Cm1 As New SqlClient.SqlCommand(SQL, Cn, lTrans)
               Cm = Cm1
            Case ClsSQL.EnumDBType.OLEDB
               Dim Cm1 As New OleDb.OleDbCommand(SQL, Cn, lTrans)
               Cm = Cm1
            Case ClsSQL.EnumDBType.ODBC
               Dim Cm1 As New Odbc.OdbcCommand(SQL, Cn, lTrans)
               Cm = Cm1
            Case ClsSQL.EnumDBType.Oracle
            Case ClsSQL.EnumDBType.MySQL
         End Select
      End If
      DrItens = Cm.ExecuteReader
      Dim Retorno As Object
      If DrItens.Read Then
         Retorno = DrItens.Item(0)
      Else
         Retorno = Nothing
      End If
      DrItens.Close()
      Return Retorno
   End Function

   Private Function Mont_SQL_Exec_Procedure() As String
      Dim SQLReturn As New System.Text.StringBuilder("Exec ")
      SQLReturn.Append(str_Procedure_Name)
      SQLReturn.Append(" ")
      Dim obj_Parameter As Cls_Procedure_Parameter
      For i As Short = 0 To obj_Collection_Parameters.Count - 1
         obj_Parameter = obj_Collection_Parameters.Item(i)
         If i > 0 Then
            SQLReturn.Append(",")
         End If
         Select Case obj_Parameter.Parameter_Type
            Case Cls_Procedure_Parameter.EnumType.eText
               SQLReturn.Append("'")
               SQLReturn.Append(obj_Parameter.Parameter_Value)
               SQLReturn.Append("'")
            Case Cls_Procedure_Parameter.EnumType.eFloat
               SQLReturn.Append(obj_Parameter.Parameter_Value.ToString.Replace(",", "."))
            Case Cls_Procedure_Parameter.EnumType.eDateTime
               SQLReturn.Append("'")
               SQLReturn.Append(Format(CType(obj_Parameter.Parameter_Value, DateTime), "yyyy/MM/dd hh:mm:ss"))
               SQLReturn.Append("'")
            Case Cls_Procedure_Parameter.EnumType.eInteger
               SQLReturn.Append(obj_Parameter.Parameter_Value.ToString)
            Case Cls_Procedure_Parameter.EnumType.eTime
               SQLReturn.Append("'")
               SQLReturn.Append(Format(CType(obj_Parameter.Parameter_Value, DateTime), "hh:mm:ss"))
               SQLReturn.Append("'")
            Case Cls_Procedure_Parameter.EnumType.eDate
               SQLReturn.Append("'")
               SQLReturn.Append(Format(CType(obj_Parameter.Parameter_Value, DateTime), "yyyy/MM/dd"))
               SQLReturn.Append("'")
            Case Cls_Procedure_Parameter.EnumType.eBool
               If CType(obj_Parameter.Parameter_Value, Boolean) = True Then
                  SQLReturn.Append("1")
               Else
                  SQLReturn.Append("0")
               End If
         End Select
      Next
      Return SQLReturn.ToString()
   End Function

   Public Property Procedure_Type() As e_Procedure_Type
      Get
         Return int_Procedure_Type
      End Get
      Set(ByVal value As e_Procedure_Type)
         int_Procedure_Type = value
      End Set
   End Property

   Public Property Procedure_Name() As String
      Get
         Return str_Procedure_Name
      End Get
      Set(ByVal value As String)
         str_Procedure_Name = value
      End Set
   End Property

   Public Property Object_Name() As String
      Get
         Return obj_Object_Name
      End Get
      Set(ByVal value As String)
         obj_Object_Name = value
      End Set
   End Property

   Public ReadOnly Property RegAffect() As Integer
      Get
         Return int_RegAffect
      End Get
   End Property

   Public ReadOnly Property Procedure_Return_Direct_Value() As Object
      Get
         Return obj_Procedure_Return_Direct_Value
      End Get
   End Property

   Public ReadOnly Property Procedure_Return_Rows() As ArrayList
      Get
         Return obj_Procedure_Return_Rows
      End Get
   End Property

   Friend ReadOnly Property Collection_Fields() As Cls_Collection_Procedure_Field
      Get
         Return obj_Collection_Fields
      End Get
   End Property

   Friend ReadOnly Property Collection_Parameters() As Cls_Collection_Procedure_Parameter
      Get
         Return obj_Collection_Parameters
      End Get
   End Property

End Class
