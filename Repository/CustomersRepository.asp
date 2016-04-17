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
            'Response.Write(count)
            set customers(count) = new Customer
            customers(count).SetIdCustomer = customersRecordSet("CustomerID").Value
            customers(count).SetCompanyName = customersRecordSet("CompanyName").Value
            customers(count).SetContactName = customersRecordSet("ContactName").Value
            customers(count).SetCity = customersRecordSet("City").Value
            customers(count).SetPhone = customersRecordSet("Phone").Value
            'Response.Write(customers(count).GetContactName())
            count = count + 1
            customersRecordSet.MoveNext()
        Loop
        customersRecordSet.Close()
        GetCustomers = customers

        'Do While NOT customersRecordSet.EOF
        '    Response.Write("<br>" & customers(count).GetIdCustomer() & " " & _
        '                   customers(count).GetCompanyName() & " " & _
        '                   customers(count).GetContactName() & " " & _
        '                   customers(count).GetCity() & " " & _ 
        '                   customers(count).GetPhone() & "<br>")
        '    count = count - 1
        '    customersRecordSet.MoveNext()
        'Loop
        
        'customersRecordSet.Close()
    End Function
%>