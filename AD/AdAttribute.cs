using System;
using System.Collections.Generic;
using System.Text;

namespace Hyperfish.ImportExport.AD
{
    public class AdAttribute
    {
        public AdAttribute(string name)
        {
            Name = name;
        }

        public string Name { get; set; }

        public bool IsMultiValued { get; set; } = false;

    }
}
