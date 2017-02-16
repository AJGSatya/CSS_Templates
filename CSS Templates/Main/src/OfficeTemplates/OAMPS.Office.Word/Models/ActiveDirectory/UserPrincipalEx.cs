using System.DirectoryServices.AccountManagement;

namespace OAMPS.Office.Word.Models.ActiveDirectory
{
    [DirectoryRdnPrefix("CN")]
    [DirectoryObjectClass("Person")]
    public class UserPrincipalEx : UserPrincipal
    {
        // Inplement the constructor using the base class constructor. 
        private UserPrincipalExSearchFilter _searchFilter;

        public UserPrincipalEx(PrincipalContext context)
            : base(context)
        {
        }

        // Implement the constructor with initialization parameters.    
        public UserPrincipalEx(PrincipalContext context,
            string samAccountName,
            string password,
            bool enabled)
            : base(context, samAccountName, password, enabled)
        {
        }

        public new UserPrincipalExSearchFilter AdvancedSearchFilter
        {
            get { return _searchFilter ?? (_searchFilter = new UserPrincipalExSearchFilter(this)); }
        }

        [DirectoryProperty("title")]
        public string Title
        {
            get
            {
                if (ExtensionGet("title").Length != 1)
                    return string.Empty;

                return (string) ExtensionGet("title")[0];
            }
            set { ExtensionSet("title", value); }
        }

        [DirectoryProperty("mobile")]
        public string Mobile
        {
            get
            {
                if (ExtensionGet("mobile").Length != 1)
                    return string.Empty;

                return (string) ExtensionGet("mobile")[0];
            }
            set { ExtensionSet("mobile", value); }
        }

        [DirectoryProperty("company")]
        public string Company
        {
            get
            {
                if (ExtensionGet("company").Length != 1)
                    return string.Empty;

                return (string) ExtensionGet("company")[0];
            }
            set { ExtensionSet("company", value); }
        }

        [DirectoryProperty("department")]
        public string Branch
        {
            get
            {
                if (ExtensionGet("department").Length != 1)
                    return string.Empty;

                return (string) ExtensionGet("department")[0];
            }
            set { ExtensionSet("department", value); }
        }

        [DirectoryProperty("Othertelephone")]
        public string BranchPhone
        {
            get
            {
                if (ExtensionGet("Othertelephone").Length != 1)
                    return string.Empty;

                return (string) ExtensionGet("Othertelephone")[0];
            }
            set { ExtensionSet("Othertelephone", value); }
        }


        [DirectoryProperty("streetAddress")]
        public string BranchAddressLine1
        {
            get
            {
                if (ExtensionGet("streetAddress").Length != 1)
                    return string.Empty;

                return (string) ExtensionGet("streetAddress")[0];
            }
            set { ExtensionSet("streetAddress", value); }
        }

        [DirectoryProperty("st")]
        public string State
        {
            get
            {
                if (ExtensionGet("st").Length != 1)
                    return string.Empty;

                return (string) ExtensionGet("st")[0];
            }
            set { ExtensionSet("st", value); }
        }

        [DirectoryProperty("co")]
        public string Country
        {
            get
            {
                if (ExtensionGet("co").Length != 1)
                    return string.Empty;

                return (string) ExtensionGet("co")[0];
            }
            set { ExtensionSet("co", value); }
        }

        [DirectoryProperty("l")]
        public string Suburb
        {
            get
            {
                if (ExtensionGet("l").Length != 1)
                    return string.Empty;

                return (string) ExtensionGet("l")[0];
            }
            set { ExtensionSet("l", value); }
        }


        [DirectoryProperty("facsimileTelephoneNumber")]
        public string Fax
        {
            get
            {
                if (ExtensionGet("facsimileTelephoneNumber").Length != 1)
                    return string.Empty;

                return (string) ExtensionGet("facsimileTelephoneNumber")[0];
            }
            set { ExtensionSet("facsimileTelephoneNumber", value); }
        }


        [DirectoryProperty("postalCode")]
        public string PostalCode
        {
            get
            {
                if (ExtensionGet("postalCode").Length != 1)
                    return string.Empty;

                return (string) ExtensionGet("postalCode")[0];
            }
            set { ExtensionSet("postalCode", value); }
        }

        //TODO find one with valid postal address and check the ad property
        [DirectoryProperty("postOfficeBox")]
        public string BranchPostalAddress
        {
            get
            {
                if (ExtensionGet("postOfficeBox").Length != 1)
                    return string.Empty;

                return (string) ExtensionGet("postOfficeBox")[0];
            }
            set { ExtensionSet("postOfficeBox", value); }
        }


        [DirectoryProperty("l")]
        public string City
        {
            get
            {
                if (ExtensionGet("l").Length != 1)
                    return string.Empty;

                return (string) ExtensionGet("l")[0];
            }
            set { ExtensionSet("l", value); }
        }


        public string BranchAddressLine2
        {
            get { return string.Format(@"{0} {1} {2}", City, State, PostalCode); }
        }


        // Implement the overloaded search method FindByIdentity.
        public new static UserPrincipalEx FindByIdentity(PrincipalContext context, string identityValue)
        {
            return (UserPrincipalEx) FindByIdentityWithType(context, typeof (UserPrincipalEx), identityValue);
        }

        // Implement the overloaded search method FindByIdentity. 
        public new static UserPrincipalEx FindByIdentity(PrincipalContext context, IdentityType identityType,
            string identityValue)
        {
            return
                (UserPrincipalEx) FindByIdentityWithType(context, typeof (UserPrincipalEx), identityType, identityValue);
        }
    }
}