using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Hyperfish.ImportExport.Helpers;
using Newtonsoft.Json;

namespace Hyperfish.ImportExport.AD
{
    public class AdConnectionCredentials
    {
        public AdConnectionCredentials(string username, SecureString password, string domain)
        {
            Username = username;
            Domain = domain;
            Password = password;
        }

        public string Username { get; set; }

        public SecureString Password { get; set; }

        public string Domain { get; set; }
    }
    
}