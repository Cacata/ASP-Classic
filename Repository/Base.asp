<%      
    Function GetConnection()
        'define the connection string
        'ConnSTring = "Provider=SQLOLEDB; Data Source = (local); Initial Catalog = Northwind; Integrated Security=SSPI;"
        ConnSTring = "DRIVER={SQL Server};Provider=MSDASQL.1;SERVER=(local); DATABASE=Northwind"

        'create an instance of the ADO connection objects
        Set Connection = Server.CreateObject("ADODB.Connection")
            
        'Open the connection to the database
        Connection.Open ConnString

        'Return connection
        Set GetConnection = Connection
    End Function



    'This function make a Select statement!
    Function GetQuery(sqlstmt)
        'Create the ADO Recordset
        Set conn = GetConnection()
        Set rs = server.createobject("ADODB.Recordset")
        rs.Open sqlstmt,conn
        
        'Return the resultant recordset
        Set GetQuery = rs
    End Function



    'This function make insert to data base
    Function MakeInsert(sqlstmt)
        Set conn = GetConnection()
        'Open the connection to the database
        conn.Execute sqlstmt,recaffected

        
        if err <> 0 then
            MakeInsert = err.Source & " " & err.Description & " " & err.HelpContext
        else
            MakeInsert = recaffected & " record added... " & err
        end if

        conn.close
    End Function



    'This function update one customer in database
    Function MakeUpdate(sqlstmt)
        Set conn = GetConnection()

        conn.Execute sqlstmt

        if err <> 0 then
            MakeUpdate = err.Source & " " & err.Description & " " & err.HelpContext
        else
            MakeUpdate = recaffected & " record added... " & err
        end if

        conn.close
    End Function



    'This function delete one customer in database
    Function DeleteOne(sqlstmt)
        Set conn = GetConnection()
        conn.Execute sqlstmt

        if err <> 0 then
            DeleteOne = err.Source & " " & err.Description & " " & err.HelpContext
        else
            DeleteOne = recaffected & " record added... " & err
        end if

        conn.close
    End Function




    'This function search one record of data base
    Function SearchRecord(sqlstmt)
        'Create the ADO objects
        Set rs = server.createobject("ADODB.Recordset")
        Set conn = GetConnection()
        rs.Open sqlstmt,conn
        
        'Return the resultant recordset 809-699-2621 Juan Peña
        Set SearchRecord = rs
    End Function
%>