<!--#include virtual="/Repository/CustomersRepository.asp"-->
<!--#include virtual="/lib/JSON_2.0.4.asp"-->
<%
    'Actions Availables
    Dim action   
    set action = Request.QueryString("action")    
    
    Dim repository, customers
    set repository = new CustomersRepository   
    
    Dim member
    set member = jsObject()
     
    'Call Action
    member("action") = action
    
    select case LCase(action)
    case "list"
        GetList()
   
    case "save"
    if Request.ServerVariables("REQUEST_METHOD")= "POST" then
        SaveCustomer()
    end if
    case "delete"
    if Request.ServerVariables("REQUEST_METHOD")= "POST" then
        DeleteCustomer()
    end if
    case "update"

   end select
    'Function
    'List
    Function GetList()
        customers = repository.GetCustomers()
        dim customer, i, total
        if isArray(customers) Then
            total = UBound(customers)       
            for each customer in customers
                if Not i = total then
                    Response.Write("<tr>")       
                    Response.Write("<td>"+ customer.GetIdCustomer + "</td>")
                    Response.Write("<td>"+ customer.GetCompanyName + "</td>")
                    Response.Write("<td>"+ customer.GetContactName + "</td>")
                    Response.Write("<td>"+ customer.GetCity + "</td>")
                    Response.Write("<td>"+ customer.GetPhone + "</td>")
                    Response.Write("<td><button class='btn btn-danger' onclick='removeCustom(this)' data-custom-id="+customer.GetIdCustomer+">Delete</button>      <button class='btn btn-warning'>Update</button></td>")
                    Response.Write("</tr>")
                End if
                i = i + 1
            next
            
        End If
    End Function

    'Save Customer
    Function SaveCustomer()
      Dim customer 
      set customer = new Customer
        '(id, company, name, city, phone)
        customer.SetCompanyName = Request.Form("CustomerId")
        customer.SetContactName = Request.Form("CompanyName")
        customer.SetContactName = Request.Form("ContactName")
        customer.SetContactName = Request.Form("CityName")
        customer.SetContactName = Request.Form("Phone")
        'repository.AddCustomer(customer)
        member("status") = "true"
        member.Flush
    End Function

    Function DeleteCustomer()
      Response.Write("<br/> Delete: " + Request.Form("id"))
      'repository.DeleteCustomer(Request.Form("id"))
    End Function
 %>