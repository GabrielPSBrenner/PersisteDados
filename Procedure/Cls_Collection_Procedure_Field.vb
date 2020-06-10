Friend Class Cls_Collection_Procedure_Field
   Inherits Collections.CollectionBase

   Public Sub New()
      MyBase.New()
   End Sub

   Default Public ReadOnly Property Item(ByVal index As Int32) As Cls_Procedure_Field
      Get
         Return CType(List.Item(index), Cls_Procedure_Field)
      End Get
   End Property

   Public Function Add(ByVal Item As Cls_Procedure_Field) As Integer
      Return List.Add(Item)
   End Function

   Public Sub Remove(ByVal Item As Cls_Procedure_Field)
      List.Remove(Item)
   End Sub

   Public Function IndexOf(ByVal value As Cls_Procedure_Field) As Integer
      Return List.IndexOf(value)
   End Function

   Public Sub Insert(ByVal index As Integer, ByVal value As Cls_Procedure_Field)
      List.Insert(index, value)
   End Sub
End Class
