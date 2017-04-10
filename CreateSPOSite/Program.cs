
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

//out usings
using Microsoft.SharePoint.Client;
using Microsoft.Online.SharePoint.TenantAdministration;
using System.Security;

//

namespace CreateSPOSite
{
    class Program
    {
        static void Main(string[] args)
        {
            var SPOTenentURL = "https://contoso005-admin.sharepoint.com/";
            var AADUsername = "admin@contoso005.onmicrosoft.com";
            var AADPassword = "Welcome@1234";
            
            //new url WITHOUT TRAILING '/'
            var NewSiteCollectionURL = "https://contoso005.sharepoint.com/sites/sitewithcode";
            var NewSiteCollectionTitle = "Make with CODE";
            var NewSiteCollectionTemplate = "STS#0";
            var NewSiteCollectionPrimaryAdministrator = "vidya.nikam@contoso005.onmicrosoft.com";
            var NewSiteCollectionSecondaryAdministrator = "";
            

            using (ClientContext ctx = new ClientContext(SPOTenentURL))
            {
                //set cred
                ctx.Credentials = new SharePointOnlineCredentials(AADUsername, GetPasswordInSecureString(AADPassword));

                //create a new tenant object
                Tenant objTenant = new Tenant(ctx);

                //create object with all new site coll details
                SiteCreationProperties newSite = new SiteCreationProperties();

                newSite.Owner = NewSiteCollectionPrimaryAdministrator;
                newSite.Title = NewSiteCollectionTitle;
                newSite.Url = NewSiteCollectionURL;
                newSite.Template = NewSiteCollectionTemplate;
                newSite.UserCodeMaximumLevel = 0;

                //create the site per above settings
                SpoOperation oSpoOps = objTenant.CreateSite(newSite);

                //load the operation
                ctx.Load(oSpoOps, spo => spo.IsComplete);

                //submit the operation
                ctx.ExecuteQuery();

                //wait for operation to be completed
                while (oSpoOps.IsComplete)
                {
                    Console.WriteLine("Waiting..{0}",DateTime.Now);

                    try
                    {
                        //keep checking if the site created
                        oSpoOps.RefreshLoad();
                        ctx.ExecuteQuery();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine("Site Creation Failed, because of {0}",e.Message);
                        break;
                    }
                }

                Console.WriteLine("Site Created");
                
            }

                Console.ReadLine();
        }
        
        private static SecureString GetPasswordInSecureString(string strPassword)
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
