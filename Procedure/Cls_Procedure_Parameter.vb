Friend Class Cls_Procedure_Parameter

   Public Enum EnumType 'Enumerador que identifica o tipo de campo 
      eFloat = 0
      eInteger = 1
      eText = 2
      eDate = 3
      eTime = 4
      eBool = 6
      eDateTime = 7 'Falta identificar o tipo data/hora junto <<IMPLEMENTAR>>
   End Enum

   Private str_Parameter_Name As String   'Nome do parâmetro que será passado
   Private obj_Parameter_Value As Object  'Valor que será passado para a procedure
   Private obj_Parameter_Type As EnumType 'Tipo do parâmetro

   Public Property Parameter_Value() As Object
      Get
         Return obj_Parameter_Value
      End Get
      Set(ByVal value As Object)
         obj_Parameter_Value = value
      End Set
   End Property

   Protected Friend Property Parameter_Name() As String
      Get
         Return str_Parameter_Name
      End Get
      Set(ByVal value As String)
         str_Parameter_Name = value
      End Set
   End Property

   Protected Friend Property Parameter_Type() As EnumType
      Get
         Return obj_Parameter_Type
      End Get
      Set(ByVal value As EnumType)
         obj_Parameter_Type = value
      End Set
   End Property
End Class
