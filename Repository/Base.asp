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



    'This function make insert to data base
    Function MakeInsert(sqlstmt)
        on error resume next
        
        'Open the connection to the database 809-699-2621 Juan Peña
        conn = GetConnection()
        conn.Execute sqlstmt,recaffected

        
        if err <> 0 then
            MakeInsert = err.Source & " " & err.Description & " " & err.HelpContext
        else
            MakeInsert = recaffected & " record added... " & err
        end if

        conn.close
    End Function



    'This function make insert to data base
    Function MakeUpdate(sqlstmt)
        on error resume next
        
        'Open the connection to the database 809-699-2621 Juan Peña
        conn = GetConnection()
        conn.Execute sqlstmt,recaffected

        if err <> 0 then
            MakeInsert = err.Source & err.Description & err.HelpContext & err.helpfile & err.Raise
        else
            MakeInsert = recaffected & " record added" & err
        end if

        conn.close
    End Function




    'This function search one record of data base
    Function SearchRecord(sqlstmt)
        'Create the ADO objects
        Set rs = server.createobject("ADODB.Recordset")

        rs.Open sqlstmt,GetConnection()
        
        'Return the resultant recordset
        Set SearchRecord = rs
    End Function
%>