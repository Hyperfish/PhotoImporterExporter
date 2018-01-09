using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Hyperfish.ImportExport.AD
{
    public class AdException : ApplicationException
    {
        public AdException(string message, Exception innerException)
            : base(message, innerException)
        {
        }

        public AdException(string message)
            : base(message)
        {
        }
    }

    public class AdConnectionException : AdException
    {
        public AdConnectionException(Exception innerException)
            : base(innerException.Message, innerException)
        {
        }
    }
}
