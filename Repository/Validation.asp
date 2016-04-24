<%
    Class Validations
        'To Validate Strings
        Function ValString(val)
            If (TypeName(val) = "String") Then
                ValString = True
            Else
                ValString = False
            End If
        End Function
        
        'To Validate Integers
        Function ValInt(val)
            If (TypeName(val) = "Integer") Then
                ValInt = True
            Else
                ValInt = False
            End If
        End Function
    End Class
%>

