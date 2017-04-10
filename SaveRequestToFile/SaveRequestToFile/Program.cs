using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.Online.SharePoint.TenantAdministration;
using System.Security;


namespace SaveRequestToFile
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.Write("Enter the SiteColl URL you want to Create: ");
            string TobeCreatedSiteColURL = Console.ReadLine();
            string username = "admin@contoso005.onmicrosoft.com";
            string password = "Welcome@1234";
            string tenantname = "contoso005";
            string tenentURL = string.Format("https://{0}-admin.sharepoint.com",tenantname);

            using (ClientContext TenantClientContext = new ClientContext(tenentURL))
            {
                
                TenantClientContext.Credentials = new SharePointOnlineCredentials(username, toSecureString(password));
                Tenant oTenant = new Tenant(TenantClientContext);
                SPOSitePropertiesEnumerable SPOEnum= oTenant.GetSiteProperties(0, true);
                TenantClientContext.Load(SPOEnum);
                TenantClientContext.ExecuteQuery();

                bool status = true;

                foreach (var item in SPOEnum)
                {
                    string CurrentSiteURL = item.Url.ToLower().Trim();
                    string ToBeCreatedSiteURL = ToBeCreatedSiteURl.ToLower().Trim();
                    if (CurrentSiteURL == ToBeCreatedSiteURl)
                    {
                        status = false;
                        break;
                    }
                }


            }

            Console.ReadKey();
            
        }
        private static SecureString toSecureString(string strPassword)
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

