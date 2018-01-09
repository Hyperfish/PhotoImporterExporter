using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Hyperfish.ImportExport
{
    public class DistinguishedName 
    {
        public DistinguishedName(string dn)
        {
            Dn = dn;
        }

        public string Dn { get; set; }
        
        public override string ToString()
        {
            return this.Dn;
        }

    }

}
