<!--#Include File="Base.asp"-->
<!--#Include virtual="Model/Customers.asp"-->
<%
    Function GetCustomers()
        stmt = "Select * from Customers"
        Dim customersRecordSet = GetQuery(stmt)
        Dim Customers(customersRecordSet.RecordCount)            
    End Function
%>