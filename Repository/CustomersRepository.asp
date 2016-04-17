<!--#Include File="Base.asp"-->
<!--#Include virtual="/Model/Customers.asp"-->
<%
    Response.Write(GetCustomers())
    Function GetCustomers()
        stmt = "SELECT * FROM Customers"
        Dim customersRecordSet
        set customersRecordSet = GetQuery(stmt)
    
        Dim customers
        int count = 0
        Do While NOT customersRecordSet.EOF
            customersRecordSet.MoveNext()
            count = count + 1
        Loop

        Response.Write("<br>" & count & "<br>")
        Redim customers(count)
        customersRecordSet.MoveFirst()
        count = 0
        Do While NOT customersRecordSet.EOF
            Response.Write(count)
            set customers(count) = new Customer
            customers(count).SetIdCustomer = customersRecordSet("ContactName").Value
            Response.Write(customers(count).GetIdCustomer())
            count = count + 1
            customersRecordSet.MoveNext()
        Loop

        Do While NOT customersRecordSet.EOF
            Response.Write("<br>" & customers(i).GetIdCustomer() & "<br>")
        Loop
    End Function
%>