using System;
using System.IO;
using log4net;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;

namespace Hyperfish.ImportExport.O365
{
    public class O365Attribute
    {
        public O365Attribute(string name, bool multivalued)
        {
            Name = name;
            IsMultiValued = multivalued;
        }

        public char MultiValueSeperator { get; set; } = '|';

        public string Name { get; set; }

        public bool IsMultiValued { get; set; }
        
    }
}
