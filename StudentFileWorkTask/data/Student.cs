using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentFileWorkTask.data
{
    internal class Student
    {
        public string Surname { get; set; }
        public string Name { get; set; }
        public string Patronymic { get; set; }
        public Group? Group { get; set; } = null;

        public Student(string surname, string name, string patronymic, Group? group = null)
        {
            Surname = surname;
            Name = name;
            Patronymic = patronymic;
            Group = group;
        }
    }
}
