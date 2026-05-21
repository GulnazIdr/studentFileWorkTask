using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentFileWorkTask.data
{
    internal class StudentResultFilter
    {
        public string FilterName { get; set; }
        public List<Filter> OptionList { get; set; }
        public StudentResultFilter(string filterName, List<Filter> optionList)
        {
            FilterName = filterName;
            OptionList = optionList;
        }
    }
}
