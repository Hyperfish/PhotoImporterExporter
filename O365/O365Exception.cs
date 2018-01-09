using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Hyperfish.ImportExport.O365.Exceptions
{
    public class O365Exception : Exception
    {
        public O365Exception(string message, Exception innerException)
        {
        }

        public O365Exception(string message)
        {
        }
    }
}
