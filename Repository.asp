<!DOCTYPE html>
<html>
    <head>

    </head>
    <body>
        <!--#include file="Global.asa"-->
        <% 

            set abc = Server.CreateObject("MS.Ad.T4")

            Response.Write(abc.correr("Armando"))
        %>
    </body>
</html>