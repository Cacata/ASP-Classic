<!--#Include File="Base.asp"-->
<!--#Include virtual="/Model/Customers.asp"-->
<%
    Class CustomersRepositry
        
        'Class Constructor
        private isConstructed
        
        private sub Class_Initialize
            isConstructed = false
        end sub
        public default function construct()
            set construct = me
            isConstructed = true
        end function    
        'Fin constructor

        'Get all customers of NorthWind DB!
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
            Response.Write("<br> El Select funciona! <br>")
            customersRecordSet.Close()
            'for i = 0 to count
            '    Response.Write("<br>" & customers(i).GetIdCustomer() & " " & _
            '                   customers(i).GetCompanyName() & " " & _
            '                   customers(i).GetContactName() & " " & _
            '                   customers(i).GetCity() & " " & _ 
            '                   customers(i).GetPhone() & "<br>")
            '    if i = count -1 Then
            '        i = 100
            '    End if
            'next
        End Function
        'End Get


        'Add new Costumer
        Function AddCustomer(id, company, name, city, phone)
            stmt = "INSERT INTO Customers (customerID,companyname," _
                 & "contactname,city,phone)" _
                 & " VALUES ('" & id & "'," _
                 & "'" & company & "'," _
                 & "'" & name & "'," _
                 & "'" & city & "'," _
                 & "'" & phone & "');"
            Result = MakeInsert(stmt)
            AddCustomer = Result
        End Function
        'End Add



        'Edit Any Costumer
        Function UpdateCustomer(id, company, name, city, phone)
              sql = "UPDATE customers SET " _
                    & "companyname='" & company & "'," _
                    & "contactname='" & name & "'," _
                    & "city='" & city & "'," _
                    & "phone='" & phone & "'" _
                    & " WHERE customerID='" & id & "'"
              Response.Write(sql)
              Result = MakeUpdate(sql)
              UpdateCustomer = Result              
        End Function
        'End Edit



        'Search some Costumer
        Function SearchCustomer(id)
            stmt = "SELECT * FROM customers WHERE customerID='" & id & "'"
            Dim customersRecordSet
            set customersRecordSet = SearchRecord(stmt)

            Dim customers
            int count = 0
            Do While NOT customersRecordSet.EOF
                customersRecordSet.MoveNext()
                count = count + 1
            Loop
            
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
            Response.Write("<br> El Select funciona! <br>")
            customersRecordSet.Close()
            'for i = 0 to count
            '    Response.Write("<br>" & customers(i).GetIdCustomer() & " " & _
            '                   customers(i).GetCompanyName() & " " & _
            '                   customers(i).GetContactName() & " " & _
            '                   customers(i).GetCity() & " " & _ 
            '                   customers(i).GetPhone() & "<br>")
            '    if i = count -1 Then
            '        i = 100
            '    End if
            'next
        End Function
        'End Search
    End Class

    set this = (new CustomersRepositry)()
    'this.GetCustomers()
    Response.Write("Pruebas<br/>")
    S = this.AddCustomer("TOYKI","La Quinta","Arturo","Bábaro","809-987-6452")
    'R = this.SearchCustomer("TRAIH")
    Response.Write(S)
%>