<%
    Class Customer  
        'Properties      
        Private IdCustomer
        Private CompanyName
        Private ContactName
        Private City
        Private Phone
        
        'private isConstructed
        
        'private sub Class_Initialize
        '    isConstructed = false
        'end sub

        'public default function construct()
        '    set construct = me
        '    isConstructed = true
        'end function  


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
            ContactName = ContactName
        End Property

        'Get and Set of IdCustomer
        public Property Let SetCity(v_city)
            City = v_city
        End Property

        public Property Get GetCity()
            GetCity = City
        End Property

        'Get and Set of IdCustomer
        public Property Let SetPhone(v_phone)
            Phone = v_phone
        End Property

        public Property Get GetPhone()
            GetPhone = Phone
        End Property

    End Class
%>