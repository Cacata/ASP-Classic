<!--#include virtual="/Repository/CustomersRepository.asp"-->
<%
    'Actions Availables
    Dim action   
    set action = Request.QueryString("action")    
    Dim repository, customers
    set repository = new CustomersRepository   
     
    Response.Write("action = " & action)
    
    select case action
    case "list"
        GetList()
   
    case "save"
    if Request.ServerVariables("REQUEST_METHOD")= "POST" then
        SaveCustomer()
    end if
   end select
    'Function
    'List
    Function GetList()
        customers = repository.GetCustomers()
        dim customer, i, total
        total = ubound(customers)
        Response.Write("<table>")
        Response.Write("<tr>")          
        for each customer in customers
            if Not i = total then    
                Response.Write("<td>"+ customer.GetCompanyName + "<td>")        
            End if
            i = i + 1
        next
        Response.Write("</tr>")
        Response.Write("</table>")
    End Function

    'Save Customer
    Function SaveCustomer()
      Dim customer 
      set customer = new Customer
       customer.SetCompanyName = Request.Form("CompanyName")
       customer.SetContactName = Request.Form("ContactName")
       
    End Function
 %>