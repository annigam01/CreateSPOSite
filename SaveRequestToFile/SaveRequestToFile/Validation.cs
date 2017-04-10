using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.Online.SharePoint.TenantAdministration;

namespace SaveRequestToFile
{
    class SpoOperations
    {

        //props
        public string SPOAdminUserName { get; set; }
        public string SPOAdminPassword { get; set; }
        public string SiteURLToBeCreated { get; set; }
        private ClientContext TenantContext { get; set; }
        public string TenantURL { get; set; }
        public string TenantName { get; set; }
        public SharePointOnlineCredentials SPOCred { get; set; }


        //Constructors
        public SpoOperations(string username, string password, string Tenant)
        {
            //this contructor overload accepts the username and password to create a client context
            
            SPOAdminUserName = username;
            SPOAdminPassword = password;
            TenantName = Tenant;
            TenantURL = string.Format("https://{0}-admin.sharepoint.com", Tenant);
            TenantContext = new ClientContext(TenantURL);
            SPOCred = new SharePointOnlineCredentials(username, toSecureString(password));
            TenantContext.Credentials = SPOCred;
            
        }
        public SpoOperations(ClientContext TenantClientContext)
        {
           
            TenantContext = TenantClientContext;
            SPOAdminUserName = null;
            SPOAdminPassword = null;
            SPOCred = null;
            TenantName = null;
            TenantURL = null;
            TenantContext = null;
            

        }

        //Validate if the site already exist
        public bool isSiteExist() {

            bool status = true;
            using (TenantContext)
            {

                
                Tenant oTenant = new Tenant(TenantContext);
                SPOSitePropertiesEnumerable SPOEnum = oTenant.GetSiteProperties(0, true);
                TenantContext.Load(SPOEnum);
                TenantContext.ExecuteQuery();
                //SPOEnum.Select(obj =>obj.Url==)

                foreach (var item in SPOEnum)
                {
                    string CurrentSiteURL= item.Url.ToLower().Trim();
                    string ToBeCreatedSiteURl = this.SiteURLToBeCreated.ToLower().Trim();
                    if (CurrentSiteURL == ToBeCreatedSiteURl)
                    {
                        status = false;
                        break;
                    }
                }


            }

            return status;

        }
        private SecureString toSecureString(string strPassword)
        {
            var secureStr = new SecureString();
            if (strPassword.Length > 0)
            {
                foreach (var c in strPassword.ToCharArray()) secureStr.AppendChar(c);
            }
            return secureStr;
        }
    }
}
