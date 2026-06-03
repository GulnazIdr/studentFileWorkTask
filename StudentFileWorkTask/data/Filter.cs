using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentFileWorkTask.data
{
    internal class Filter
    {
        public string Option { get; set; }
        public bool IsChecked { get; set; }
        public Filter(string option, bool isChecked)
        {
            Option = option;
            IsChecked = isChecked;
        }
    }
}
