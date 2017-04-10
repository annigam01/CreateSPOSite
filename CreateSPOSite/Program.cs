
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
        public bool isSiteExist(string SiteURLToBeCreated, ClientContext TenantContext)
        {
            bool status = true;

            using (TenantContext)
            {
                // get tenant object
                Tenant oTenant = new Tenant(TenantContext);

                //get all SPO site property - which includes all the site collections
                SPOSitePropertiesEnumerable SPOEnum = oTenant.GetSiteProperties(0, true);

                //load the prop object
                TenantContext.Load(SPOEnum);

                //execute the query
                TenantContext.ExecuteQuery();
                
                //iterate thru to check if the url matches, if yes retun true 
                foreach (var item in SPOEnum)
                {
                    //current SPO site collection URL
                    string CurrentSiteURL = item.Url.ToLower().Trim();
                    //Site collection url you want to create
                    string ToBeCreatedSiteURl = SiteURLToBeCreated.ToLower().Trim();

                    //do the url match?
                    if (CurrentSiteURL == ToBeCreatedSiteURl)
                    {
                        status = false;
                        break;
                    }
                }


            }

            return status;

        }
        public void TurnOnExternalShareing(string SiteURL, ClientContext TenantContext)
        {
            //this method turns the "exernal sharing feature" to ON, if its already ON, nothing is done
            //get tenant object
            Tenant oTenant = new Tenant(TenantContext);
            
            //get the sitecollection property of desired site
            SiteProperties SiteCollProp = oTenant.GetSitePropertiesByUrl(SiteURL, true);
            
            //load object
            TenantContext.Load(SiteCollProp);
            
            //get data from server
            TenantContext.ExecuteQuery();

            //read the office article to understand the sharing feature then turn the flag on off as desired
            // https://support.office.com/en-us/article/Manage-external-sharing-for-your-SharePoint-Online-environment-c8a462eb-0723-4b0b-8d0a-70feafe4be85


            //select the appropriate type of external sharing flag, you enum SharingCapabilities for all options
            SiteCollProp.SharingCapability =
                Microsoft.Online.SharePoint.TenantManagement.SharingCapabilities.ExternalUserSharingOnly;

            //other sharing options as below

            //1. set the whitelisted domains here
            SiteCollProp.SharingAllowedDomainList = "enter the domains here";

            //2. OR set the black listed domains here
            SiteCollProp.SharingBlockedDomainList = "enter the domains here";

            //once all the changes are done, call update method
            SiteCollProp.Update();

            TenantContext.ExecuteQuery();
            
        }
    }
}
