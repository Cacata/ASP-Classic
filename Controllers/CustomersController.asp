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
    'member("action") = action
    
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
    case "get"
    if Request.ServerVariables("REQUEST_METHOD")= "GET" then
        GetById()
    end if
    case "update"
    UpdateCustomer()
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
                    Response.Write("<td><button class='btn btn-danger' onclick='removeCustom(this)' data-custom-id="+customer.GetIdCustomer+">Delete</button>      <button class='btn btn-warning'  onclick='getCustomerId(this)' data-custom-id="+customer.GetIdCustomer+">Editar</button></td>")
                    Response.Write("</tr>")
                End if
                i = i + 1
            next
            
        End If
    End Function

    Function GetById()
     Dim id 
     set id = Request.QueryString("Id")
     Dim customer 
     set customer = repository.SearchCustomer(id)
    'Response.Write(customer.GetCompanyName())
      member("IdCustomer") = customer.GetIdCustomer
      member("CompanyName") = customer.GetCompanyName
      member("ContactName") =  customer.GetContactName
      member("City") = customer.GetCity
      member("Phone") = customer.GetPhone
      member.Flush()
    End Function

    'Save Customer
    Function SaveCustomer()
      Dim customer 
      set customer = new Customer
        '(id, company, name, city, phone)
        customer.SetIdCustomer = Request.Form("CustomerId")
        customer.SetCompanyName = Request.Form("CompanyName")
        customer.SetContactName = Request.Form("ContactName")
        customer.SetCity = Request.Form("CityName")
        customer.SetPhone = Request.Form("Phone")
        
        'repository.AddCustomer customer.GetIdCustomer, customer.GetCompanyName(), customer.GetContactName(), customer.GetCity(), customer.GetPhone()
        call repository.AddCustomer(customer.GetIdCustomer, customer.GetCompanyName(), customer.GetContactName(), customer.GetCity(), customer.GetPhone())
        Response.ContentType = "text/json"
        member("status") = "true"
        member.Flush
    End Function

    Function UpdateCustomer()
      Dim customer 
      set customer = new Customer
        '(id, company, name, city, phone)
        customer.SetIdCustomer = Request.Form("CustomerId")
        customer.SetCompanyName = Request.Form("CompanyName")
        customer.SetContactName = Request.Form("ContactName")
        customer.SetCity = Request.Form("CityName")
        customer.SetPhone = Request.Form("Phone")
        
        'repository.AddCustomer customer.GetIdCustomer, customer.GetCompanyName(), customer.GetContactName(), customer.GetCity(), customer.GetPhone()
        call repository.UpdateCustomer(customer.GetIdCustomer, customer.GetCompanyName(), customer.GetContactName(), customer.GetCity(), customer.GetPhone())
        Response.ContentType = "text/json"
        member("status") = "true"
        member.Flush
    End Function


    Function DeleteCustomer()
      set id = Request.Form("id")
      repository.DeleteCustomer(id)
        
        Response.ContentType = "text/json"
        member("status") = "true"
        member.Flush
    End Function
 %>