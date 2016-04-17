<%      
    Sub GetConnection()
        'define the connection string, specify database driver
        ConnString="DRIVER={SQL Server};SERVER=(local); DATABASE=Northwind"

        'create an instance of the ADO connection objects
        Set Connection = Server.CreateObject("ADODB.Connection")
            
        'Open the connection to the database
        Connection.Open ConnString

        'Return connection
        Set GetConnection = Connection
    End Sub


    'Function RunSQLReturnRS(sqlstmt, params())
    Function GetQuery(sqlstmt)
        On Error Resume next

        'Create the ADO objects
        Dim rs , cmd
        Set rs = server.createobject("ADODB.Recordset")
        Set cmd = server.createobject("ADODB.Command")

        'Init the ADO objects  & the stored proc parameters
        cmd.ActiveConnection = GetConnection()
        cmd.CommandText = sqlstmt
        cmd.CommandType = adCmdText
        cmd.CommandTimeout = 900 

        'propietary function that put params in the cmd
        'collectParams cmd, params

        'Execute the query for readonly
        rs.CursorLocation = adUseClient
        rs.Open cmd, , adOpenForwardOnly, adLockReadOnly

        If err.number > 0 then
            BuildErrorMessage()
            exit function
        end if

        'Disconnect the recordset
        Set cmd.ActiveConnection = Nothing
        Set cmd = Nothing
        Set rs.ActiveConnection = Nothing

        'Return the resultant recordset
        Set GetQuery = rs
    End Function
%>