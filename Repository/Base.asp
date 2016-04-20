<%      
    conn = GetConnection()

    Function GetConnection()
        'define the connection string
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
        'Create the ADO Recordset
        Set rs = server.createobject("ADODB.Recordset")
        rs.Open sqlstmt,conn
        
        'Return the resultant recordset
        Set GetQuery = rs
    End Function



    'This function make insert to data base
    Function MakeInsert(sqlstmt)
        
        'Open the connection to the database
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

        conn.Execute sql

        if err <> 0 then
            MakeUpdate = err.Source & " " & err.Description & " " & err.HelpContext
        else
            MakeUpdate = recaffected & " record added... " & err
        end if

        conn.close
    End Function




    'This function search one record of data base
    Function SearchRecord(sqlstmt)
        'Create the ADO objects
        Set rs = server.createobject("ADODB.Recordset")

        rs.Open sqlstmt,conn
        
        'Return the resultant recordset 809-699-2621 Juan Peña
        Set SearchRecord = rs
    End Function
%>