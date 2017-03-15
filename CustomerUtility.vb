Imports System.Dynamic
Imports System.Reflection
Imports System.Web.Script.Serialization
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters

Public Class CustomerUtility

    Public Shared Sub AddProperty(expando As ExpandoObject, propertyName As String, propertyValue As Object)
        ' ExpandoObject supports IDictionary so we can extend it like this
        Dim expandoDict = TryCast(expando, IDictionary(Of String, Object))
        If expandoDict.ContainsKey(propertyName) Then
            expandoDict(propertyName) = propertyValue
        Else
            expandoDict.Add(propertyName, propertyValue)
        End If
    End Sub
    Public Shared Sub RemoveProperty(expando As ExpandoObject, propertyName As String)
        ' ExpandoObject supports IDictionary so we can extend it like this
        Dim expandoDict = TryCast(expando, IDictionary(Of String, Object))
        expandoDict.Remove(propertyName)
    End Sub


    Public Shared Function ConvertToFlatJson(cust As Customer) As ExpandoObject
        Dim expConverter = New ExpandoObjectConverter()
        Dim mycust As ExpandoObject = JsonConvert.DeserializeObject(Of ExpandoObject)(JsonConvert.SerializeObject(cust), expConverter)

        For Each field In cust.CustomerFieldList
            AddProperty(mycust, field.FieldName, field.FieldValue)
        Next

        RemoveProperty(mycust, "CustomerFieldList")

        Return mycust
    End Function
End Class

Public Class PropertyUtils
    Private Sub New()
    End Sub



    ''' --------------------------------------------------------------------
    ''' <summary>
    ''' Determine if a property exists in an object
    ''' </summary>
    ''' <param name="propertyName">Name of the property </param>
    ''' <param name="srcObject">the object to inspect</param>
    ''' <returns>true if the property exists, false otherwise</returns>
    ''' <exception cref="ArgumentNullException">if srcObject is null</exception>
    ''' <exception cref="ArgumentException">if propertName is empty or null </exception>
    ''' --------------------------------------------------------------------
    Public Shared Function Exists(propertyName As String, srcObject As Object) As Boolean
        If srcObject Is Nothing Then
            Throw New System.ArgumentNullException("srcObject")
        End If

        If (propertyName Is Nothing) OrElse (propertyName = [String].Empty) OrElse (propertyName.Length = 0) Then
            Throw New System.ArgumentException("Property name cannot be empty or null.")
        End If

        Dim propInfoSrcObj As PropertyInfo = srcObject.[GetType]().GetProperty(propertyName)

        Return (propInfoSrcObj IsNot Nothing)
    End Function


    ''' --------------------------------------------------------------------
    ''' <summary>
    ''' Determine if a property exists in an object
    ''' </summary>
    ''' <param name="propertyName">Name of the property </param>
    ''' <param name="srcObject">the object to inspect</param>
    ''' <param name="ignoreCase">ignore case sensitivity</param>
    ''' <returns>true if the property exists, false otherwise</returns>
    ''' <exception cref="ArgumentNullException">if srcObject is null</exception>
    ''' <exception cref="ArgumentException">if propertName is empty or null </exception>
    ''' --------------------------------------------------------------------
    Public Shared Function Exists(propertyName As String, srcObject As Object, ignoreCase As Boolean) As Boolean
        If Not ignoreCase Then
            Return Exists(propertyName, srcObject)
        End If

        If srcObject Is Nothing Then
            Throw New System.ArgumentNullException("srcObject")
        End If

        If (propertyName Is Nothing) OrElse (propertyName = [String].Empty) OrElse (propertyName.Length = 0) Then
            Throw New System.ArgumentException("Property name cannot be empty or null.")
        End If


        Dim propertyInfos As PropertyInfo() = srcObject.[GetType]().GetProperties()

        propertyName = propertyName.ToLower()
        For Each propInfo As PropertyInfo In propertyInfos
            If propInfo.Name.ToLower().Equals(propertyName) Then
                Return True
            End If
        Next
        Return False
    End Function
End Class




