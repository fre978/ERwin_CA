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
        If model.Properties.HasProperty(prop) Then
            model.Properties(prop).Value = value
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary>
    ''' Assign an INTEGER value to a model
    ''' </summary>
    ''' <param name="model"></param>
    ''' <param name="prop"></param>
    ''' <param name="value"></param>
    ''' <returns></returns>
    Public Function AssignToObjModel(ByRef model As SCAPI.ModelObject,
                                     prop As String, value As Integer) As Boolean
        If model.Properties.HasProperty(prop) Then
            model.Properties(prop).Value = value
            Return True
        Else
            Return False
        End If
    End Function



End Class
