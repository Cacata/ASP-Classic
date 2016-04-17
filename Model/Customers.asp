<%
    Class Customer  
        'Properties      
        Private _IdCustomer
        Private _CompanyName
        Private _ContactName
        Private _City
        Private _Phone
        
        'Get and Set of _IdCustomer
        public Property Let SetIdCustomer(id)
            _IdCustomer = id
        End Property

        public Property Get GetIdCustomer()
            GetIdCustomer = _IdCustomer
        End Property

        'Get and Set of _CompanyName
        public Property Let SetCompanyName(company)
            _CompanyName = company
        End Property

        public Property Get GetCompanyName()
            GetCompanyName = _CompanyName
        End Property

        'Get and Set of _ContactName
        public Property Let SetContactName(contact)
            _ContactName = contact
        End Property

        public Property Get GetCompanyName()
            GetCompanyName = _ContactName
        End Property

        'Get and Set of _IdCustomer
        public Property Let SetCity(city)
            _City = city
        End Property

        public Property Get GetCity()
            GetCity = _City
        End Property

        'Get and Set of _IdCustomer
        public Property Let SetPhone(phone)
            _Phone = phone
        End Property

        public Property Get GetPhone()
            GetPhone = _Phone
        End Property

    End Class
%>