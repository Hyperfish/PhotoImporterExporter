using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Hyperfish.ImportExport;

namespace Hyperfish.ImportExport
{
    public class UpnIdentifier : IAdIdentifier
    {
        public UpnIdentifier(string upn)
        {
            Upn = upn;
        }
        
        public string Upn { get; set; }
        
        public override string ToString()
        {
            return Upn;
        }
    }
}
