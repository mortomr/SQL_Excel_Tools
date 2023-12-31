'simple vba UDFs for formatting [fields] or 'values' inside Excel for use in SQL

Function SQLFields(ParamArray args() As Variant) As String
    Dim cell As Range
    Dim result As String
    Dim rng As Range
    
    result = ""
    
    For Each arg In args
        If TypeName(arg) = "Range" Then
            Set rng = arg
            For Each cell In rng
                result = result & "[" & Trim(cell.Value) & "],"
            Next cell
        ElseIf TypeName(arg) = "String" Then
            Set rng = Range(arg)
            For Each cell In rng
                result = result & "[" & Trim(cell.Value) & "],"
            Next cell
        End If
    Next arg
    
    ' Remove the trailing comma
    If Len(result) > 0 Then
        result = Left(result, Len(result) - 1)
    End If
    
    SQLFields = result
End Function

Function SQLText(ParamArray args() As Variant) As String
    Dim cell As Range
    Dim result As String
    Dim rng As Range
    
    result = ""
    
    For Each arg In args
        If TypeName(arg) = "Range" Then
            Set rng = arg
            For Each cell In rng
                result = result & "'" & Trim(cell.Value) & "',"
            Next cell
        ElseIf TypeName(arg) = "String" Then
            Set rng = Range(arg)
            For Each cell In rng
                result = result & "'" & Trim(cell.Value) & "',"
            Next cell
        End If
    Next arg
    
    ' Remove the trailing comma
    If Len(result) > 0 Then
        result = Left(result, Len(result) - 1)
    End If
    
    SQLText = result
End Function
