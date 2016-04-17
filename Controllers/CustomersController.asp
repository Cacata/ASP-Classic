<!--#include virtual="/Repository/CustomersRepository.asp"-->
<%
    'Actions Availables
    Dim action 
    set action = Request.QueryString("action")    

    select case action
    case "list"
        GetList()
   
    case "save"
    if Request.ServerVariables("REQUEST_METHOD")= "POST"
        SaveCustomer()
    end if
   end select
    'Function
    'List
    Function GetList()
       Dim customers
       Set customers = GetCustomers()

       for i = 0 to UBound(customers)
           Dim customer
          set customer = customers(i)
          Response.Write("<tr>")
          Response.Write("<td>"&customer.GetCompanyName()&"</td>")
          Response.Write("</tr>")
       next
    End Function

    'Save Customer
    Function SaveCustomer()
      Dim customer 
      set customer = new Customer
       customer.SetCompanyName = Request.Form("CompanyName")
       customer.SetContactName = Request.Form("ContactName")

    End Function
 %>