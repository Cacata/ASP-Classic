<%      
    Function GetConnection()
        'define the connection string, specify database driver
        'ConnString="DRIVER={SQL Server};SERVER=(local); DATABASE=Northwind"
        ConnSTring = "Provider=SQLOLEDB; Data Source = (local); Initial Catalog = Northwind; Integrated Security=SSPI;"

        'create an instance of the ADO connection objects
        Set Connection = Server.CreateObject("ADODB.Connection")
            
        'Open the connection to the database
        Connection.Open ConnString

        'Return connection
        Set GetConnection = Connection
    End Function

    'This function make a Select statement!
    Function GetQuery(sqlstmt)
        'Create the ADO objects
        Set rs = server.createobject("ADODB.Recordset")

        rs.Open sqlstmt,GetConnection()
        
        'Return the resultant recordset
        Set GetQuery = rs
    End Function
%>