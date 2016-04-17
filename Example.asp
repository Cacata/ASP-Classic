<!DOCTYPE html>
<html>
    <head>
        <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/css/bootstrap.min.css" 
            integrity="sha384-1q8mTJOASx8j1Au+a5WDVnPi2lkFfwwEAa8hDDdjZlpLegxhjVME1fgjWPGmkzs7" 
            crossorigin="anonymous">
        <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.6/js/bootstrap.min.js" 
            integrity="sha384-0mSbJDEHialfmuBBQP6A4Qrprq5OVfW37PRR3j5ELqxss1yVqOtnepnHVP9aJ7xS" 
            crossorigin="anonymous" type="text/javascript"></script>
        <script type="text/javascript" src="https://code.jquery.com/jquery-2.2.3.min.js"></script>
    </head>
    <body>
        <% 
        'declare the variables 
        Dim Connection
        Dim ConnString
        Dim Recordset
        Dim SQL

        'define the connection string, specify database driver
        ConnString="DRIVER={SQL Server};SERVER=(local); DATABASE=Northwind"

        'declare the SQL statement that will query the database
        SQL = "SELECT * FROM Customers"

        'create an instance of the ADO connection and recordset objects
        Set Connection = Server.CreateObject("ADODB.Connection")
        Set Recordset = Server.CreateObject("ADODB.Recordset")
            
        'Open the connection to the database
        Connection.Open ConnString

        'Open the recordset object executing the SQL statement and return records 
        Recordset.Open SQL,Connection

        %>
        <div style="height: 600px; width: 500px; overflow-y: scroll; margin: 20px;">
            <table class="table table-bordered">
            <tr>
                <td style="font-weight: bold">ContactName</td>
                <td style="font-weight: bold">CompanyName</td>
                <td style="font-weight: bold">ContactTitle</td>
            </tr>
            <%
                'first of all determine whether there are any records 
            If Recordset.EOF Then 
            Response.Write("No records returned.") 
            Else 
            'if there are records then loop through the fields 
                Do While NOT Recordset.EOF
                    Response.write("<tr>")
                    Response.write("<td>" & Recordset("ContactName") & "</td>")
                    Response.write("<td>" & Recordset("CompanyName") & "</td>")
                    Response.write("<td>" & Recordset("ContactTitle") & "</td>")
                    Response.write("</tr>")
                    Recordset.MoveNext     
                    Loop
                End If
                 %>
        </table>
        </div>
            
        <div style="width: 60%; margin-left: 20px;">
            <button class="btn btn-default" title="Agregar">Agregar</button>
            <button class="btn btn-default" title="Editar">Editar</button>
            <button class="btn btn-default" title="Eliminar">Eliminar</button>
            <button class="btn btn-default" title="Eliminar">Ordenar Lista</button>
        </div>
        <!--End VBScripts-->
    </body>
</html>