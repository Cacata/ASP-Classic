<!--#Include File="Base.asp"-->
<!--#Include File="Validation.asp"-->
<!--#Include virtual="/Model/Customers.asp"-->
<%
    Class CustomersRepository
        
        'Class Constructor
        private isConstructed     
        Dim validate
        Dim ObjectFinded
        
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
            customersRecordSet.Close()
            GetCustomers = customers
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
            'Validate values
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
            set validate = new Validation

            'Validate values
            if validate.ValString(id) = false then
                UpdateCustomer = "Id "& id &" value isn´t string"
            elseif validate.ValString(company) = false then
                UpdateCustomer = "company "& company &" value isn´t string"
            elseif validate.ValString(name) = false then
                UpdateCustomer = "name "& name &" value isn´t string"
            elseif validate.ValString(city) = false then
                UpdateCustomer = "city "& city &" value isn´t string"
            elseif validate.ValString(phone) = false then
                UpdateCustomer = "phone "& phone &" value isn´t string"
            else
                sql = "UPDATE customers SET " _
                    & "companyname='" & company & "'," _
                    & "contactname='" & name & "'," _
                    & "city='" & city & "'," _
                    & "phone='" & phone & "'" _
                    & " WHERE customerID='" & id & "'"
                Result = MakeUpdate(sql)
                UpdateCustomer = Result 
            End if      
        End Function
        'End Edit

        'Delete Any Costumer
        Function DeleteCustomer(id)
            set validate = new Validation
            if validate.ValString(id) = false then
                DeleteCustomer = "Id "& id &" value isn´t string"
            else
                sql="DELETE FROM customers WHERE customerID='" & id & "'"
                Result = DeleteOne(sql)
                DeleteCustomer = Result
            End if            
        End Function    
        'End Delete

        'Search some Costumer
        Function SearchCustomer(id)
                Dim stmt 
                stmt = "SELECT * FROM customers WHERE customerID='" & id & "'"
                Dim customersRecordSet
                set customersRecordSet = SearchRecord(stmt) 
                Dim customers
                set customers = new Customer

                customers.SetIdCustomer = customersRecordSet("CustomerID").Value
                customers.SetCompanyName = customersRecordSet("CompanyName").Value
                customers.SetContactName = customersRecordSet("ContactName").Value
                customers.SetCity = customersRecordSet("City").Value
                customers.SetPhone = customersRecordSet("Phone").Value
                
                customersRecordSet.Close()
                set SearchCustomer = customers        
            End if
        End Function
    End Class
%>