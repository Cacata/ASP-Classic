<%
    Class Myclass()        
        Private myname = "Armando"
        
        public Property Let NewName(name)
            myname = name
        End Property

        public Property Get MyName()
            MyName = myname
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