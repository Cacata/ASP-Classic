<%
    Class Customer  
        'Properties      
        Private IdCustomer
        Private CompanyName
        Private ContactName
        Private City
        Private Phone
        
        'Get and Set of IdCustomer
        public Property Let SetIdCustomer(id)
            IdCustomer = id
        End Property

        public Property Get GetIdCustomer()
            GetIdCustomer = IdCustomer
        End Property

        'Get and Set of CompanyName
        public Property Let SetCompanyName(company)
            CompanyName = company
        End Property

        public Property Get GetCompanyName()
            GetCompanyName = CompanyName
        End Property

        'Get and Set of ContactName
        public Property Let SetContactName(contact)
            ContactName = contact
        End Property

        public Property Get GetContactName()
            GetCompanyName = ContactName
        End Property

        'Get and Set of IdCustomer
        public Property Let SetCity(city)
            City = city
        End Property

        public Property Get GetCity()
            GetCity = City
        End Property

        'Get and Set of IdCustomer
        public Property Let SetPhone(phone)
            Phone = phone
        End Property

        public Property Get GetPhone()
            GetPhone = Phone
        End Property

    End Class
%>