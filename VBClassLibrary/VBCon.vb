Public Class VBCon
    'Public Function VBWriteLine(ByVal text As String)
    '    Console.WriteLine(text)
    'End Function

    'Public Function AddNumber(ByVal num1 As Integer,
    'ByVal num2 As Integer) As Integer
    '    Return num1 + num2
    'End Function


    ''' <summary>
    ''' Assign a STRING value to a model
    ''' </summary>
    ''' <param name="model"></param>
    ''' <param name="prop"></param>
    ''' <param name="value"></param>
    ''' <returns></returns>
    Public Function AssignToObjModel(ByRef model As SCAPI.ModelObject,
                                     prop As String, value As String) As Boolean
        Try
            model.Properties(prop).Value = value
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function AssignToObjModelTEST(model As SCAPI.ModelObject,
                                     prop As String, value As String) As SCAPI.ModelObject
        Try
            model.Properties(prop).Value = value
            Return model
        Catch ex As Exception
            Return model
        End Try
    End Function


    ''' <summary>
    ''' Assign an INTEGER value to a model
    ''' </summary>
    ''' <param name="model"></param>
    ''' <param name="prop"></param>
    ''' <param name="value"></param>
    ''' <returns></returns>
    Public Function AssignToObjModelInt(ByRef model As SCAPI.ModelObject,
                                     prop As String, value As Integer) As Boolean
        Try
            model.Properties(prop).Value = value
            Return True
        Catch ex As Exception
            Return False
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Read property of an object
    ''' </summary>
    ''' <param name="model"></param>
    ''' <param name="prop"></param>
    ''' <param name="value"></param>
    ''' <returns></returns>
    Public Function RetrieveFromObjModel(model As SCAPI.ModelObject,
                                     prop As String, ByRef value As String) As Boolean
        Try
            If model.Properties.HasProperty(prop) Then
                value = model.Properties(prop).Value
                Return True
            Else
                Return False
            End If
        Catch
            Return False
        End Try
    End Function

    Public Function RetriveEntity(ByRef model As SCAPI.ModelObject,
                                  collection As SCAPI.ModelObjects,
                                  entityName As String) As Boolean
        Try
            model = collection.Item(entityName, "Entity")

            '            model.Properties.Add()
            Return True
        Catch exp As Exception
            Return False
        End Try
    End Function

    Public Function RetriveAttribute(ByRef model As SCAPI.ModelObject,
                                  collection As SCAPI.ModelObjects,
                                  entityName As String) As Boolean
        Try
            model = collection.Item(entityName, "Attribute")
            '            model.Properties.Add()
            Return True
        Catch exp As Exception
            Return False
        End Try
    End Function

    Public Function RetriveRelation(ByRef model As SCAPI.ModelObject,
                                  collection As SCAPI.ModelObjects,
                                  entityName As String) As Boolean
        Try
            model = collection.Item(entityName, "Relationship")
            '            model.Properties.Add()
            Return True
        Catch exp As Exception
            Return False
        End Try
    End Function

    Public Function AssignToAttributeX(ByRef model As SCAPI.ModelObject, collection As SCAPI.ModelObjects, parent As String, child As String, relation As String) As Boolean
        For Each element As SCAPI.ModelObject In collection
            If element.Properties("Name").Value = child Then
                Try
                    element.Properties("Parent_Attribute_Ref").Value = parent
                    element.Properties("Parent_Relationship_Ref").Value = relation
                    Return True
                Catch
                    Return False
                End Try
            End If
        Next
        Return False
        Throw New NotImplementedException()
    End Function

    'Public Function AssignValueToProperty(collection As SCAPI.ModelProperties, prop As String, value As String) As Boolean
    '    Try
    '        collection.Item(prop).Value = value
    '        Return True
    '    Catch
    '        Return False
    '    End Try
    'End Function

End Class
