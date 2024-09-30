' UDF is ~30% slower than native implementation... but easier to remember :)
'Native Function Example: =IF(J4<((H4-G4)/((H4-G4)+(I4-H4))),(G4+SQRT((J4*(H4-G4)*((H4-G4)+(I4-H4))))),(I4-SQRT((1-J4)*(I4-H4)*((H4-G4)+(I4-H4)))))
Function Risk3Point(LOW As Double, MID As Double, HI As Double, PCT As Double) As Double
    Dim result As Double
    
    If PCT < ((MID - LOW) / ((MID - LOW) + (HI - MID))) Then
        result = LOW + Sqr(PCT * (MID - LOW) * ((MID - LOW) + (HI - MID)))
    Else
        result = HI - Sqr((1 - PCT) * (HI - MID) * ((MID - LOW) + (HI - MID)))
    End If
    
    Risk3Point = result
End Function


'simple vba UDFs for formatting [fields] or 'values' inside Excel for use in SQL

Function SQLFields(ParamArray args() As Variant) As String
    Dim cell As Range
    Dim result As String
    Dim rng As Range
    Dim arg as Variant
    
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
    Dim arg as Variant
    
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
