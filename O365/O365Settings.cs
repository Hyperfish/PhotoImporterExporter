using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using Newtonsoft.Json.Linq;

namespace Hyperfish.ImportExport.O365
{
    public class O365Settings
    {
        private const string ProfileSiteTemplateUrl = "https://{0}-admin.sharepoint.com";
        private const string MySiteHostTemplateUrl = "https://{0}-my.sharepoint.com";
        private const string RootSiteTemplateUrl = "https://{0}.sharepoint.com";

        public const string SpoProfilePrefix = "i:0#.f|membership|";

        public string MySiteHost => string.Format(MySiteHostTemplateUrl, this.TenantName);

        public string RootSiteHost => string.Format(RootSiteTemplateUrl, this.TenantName);

        public string ExoServiceUri => "https://outlook.office365.com/EWS/Exchange.asmx";

        public string AdminUri => string.Format(ProfileSiteTemplateUrl, this.TenantName);

        public string TenantName { get; set; }

        public string Username { get; set; }

        public SecureString Password { get; set; }
        
    }
}
