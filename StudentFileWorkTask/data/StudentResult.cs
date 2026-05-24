using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentFileWorkTask.data
{
    internal class StudentResult
    {
        public Student Student { get; set; }
        public double Score { get; set; }
        public Question Question { get; set; } 
        public DateOnly? Date { get; set; } = null;

        public StudentResult(Student student, Question question, double score,  DateOnly? date = null)
        {
            Question = question;
            Score = score;
            Date = date;
            Student = student;
        }
    }
}
