Imports System.Reflection

Friend Class ClsFunctions
   Public Shared Function CreateInstance(ByVal lAssemblyName As String, ByVal lTypeName As String) As Object
      Dim MyAssembly As Assembly = Assembly.Load(lAssemblyName)
      Return (MyAssembly.CreateInstance(lTypeName, False, BindingFlags.CreateInstance, Nothing, Nothing, Nothing, Nothing))
   End Function
End Class
