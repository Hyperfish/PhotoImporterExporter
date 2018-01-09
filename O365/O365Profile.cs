using System;
using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Hyperfish.ImportExport.O365
{
    public class O365Profile
    {
        private const string PictureAttributeName = "PictureURL";

        private UpnIdentifier _upn;
        
        public O365Profile(UpnIdentifier upn)
        {
            this.Properties = new Dictionary<string, object>();
            this.Upn = upn;
        }

        public UpnIdentifier Upn
        {
            get { return _upn; }
            set { _upn = value; }
        }

        public IDictionary<string, object> Properties { get; set; }

        public Uri SpoProfilePictureUrl
        {
            get
            {
                if (this.Properties.ContainsKey(PictureAttributeName))
                    return new Uri(this.Properties[PictureAttributeName].ToString());
                else
                    return null;
            }
        }

        public Uri SpoProfilePictureUrlLarge
        {
            get
            {
                if (this.Properties.ContainsKey(PictureAttributeName))
                    return new Uri(this.Properties[PictureAttributeName].ToString()?.Replace("MThumb","LThumb"));
                else
                    return null;
            }
        }

        [JsonIgnore]
        public bool HasSpoProfilePicture => !string.IsNullOrEmpty(SpoProfilePictureUrl?.OriginalString);
    }
}
