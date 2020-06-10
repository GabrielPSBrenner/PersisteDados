Friend Class Cls_Collection_Procedure_Parameter
   Inherits Collections.CollectionBase

   Public Sub New()
      MyBase.New()
   End Sub

   Default Public ReadOnly Property Item(ByVal index As Int32) As Cls_Procedure_Parameter
      Get
         Return CType(List.Item(index), Cls_Procedure_Parameter)
      End Get
   End Property

   Public Function Add(ByVal Item As Cls_Procedure_Parameter) As Integer
      Return List.Add(Item)
   End Function

   Public Sub Remove(ByVal Item As Cls_Procedure_Parameter)
      List.Remove(Item)
   End Sub

   Public Function IndexOf(ByVal value As Cls_Procedure_Parameter) As Integer
      Return List.IndexOf(value)
   End Function

   Public Sub Insert(ByVal index As Integer, ByVal value As Cls_Procedure_Parameter)
      List.Insert(index, value)
   End Sub
End Class
