using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentFileWorkTask.data
{
    internal class Template
    {
        public bool IsGroupExists { get; set; }
        public bool IsDataExists { get; set; }
        public Template(bool isDataExists, bool isGroupExists)
        {
            IsGroupExists = isGroupExists;
            IsDataExists = isDataExists;
        }
    }
}
