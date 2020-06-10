Friend Class Cls_Procedure_Field
   Private str_Property_Name As String 'Nome da propriedade no objeto que representa o parâmetro
   Private str_Field_Name As String    'Nome do campo parametro do procedure

   Public Property Property_Name() As String
      Get
         Return str_Property_Name
      End Get
      Set(ByVal value As String)
         str_Property_Name = value
      End Set
   End Property

   Public Property Field_Name() As String
      Get
         Return str_Field_Name
      End Get
      Set(ByVal value As String)
         str_Field_Name = value
      End Set
   End Property
End Class
