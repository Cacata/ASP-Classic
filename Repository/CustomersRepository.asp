<!--#Include File="Base.asp"-->
<!--#Include virtual="/Model/Customers.asp"-->
<%
    Function GetCustomers()
        stmt = "Select * from Customers"
        Dim customersRecordSet, customers
        set customersRecordSet = GetQuery(stmt)            


    End Function
%>