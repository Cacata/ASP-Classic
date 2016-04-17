<!--#Include File="Base.asp"-->
<!--#Include virtual="/Model/Customers.asp"-->
<%
    Class CustomersRepositry
        
        'Constructor de la clase
        private isConstructed
        private name

        private sub Class_Initialize
            isConstructed = false
            name = null
        end sub
        public default function construct(pName)
            name = pName
            set construct = me
            isConstructed = true
        end function    
        'Fin constructor

        'set this = (New CustomersRepositry)("Armando")
            'Dim this = new CustomersRepositry
            'Response.Write(this.GetCustomers())

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
    End Class

    set this = (new CustomersRepositry)("Armando")
    this.GetCustomers()
%>