<%
    Class Myclass        
        Private my_name
        
        public Property Let NewName(name)
            my_name = name
        End Property

        public Property Get MyName()
            MyName = my_name
        End Property

        Sub correr(nom)
            correr = nom & "esta corriendo."
        End Sub

        Sub caminar(nom)
            caminar = nom & "esta caminando."
        End Sub
    End Class

    Set yo = new Myclass
%>
