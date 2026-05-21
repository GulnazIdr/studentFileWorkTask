using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentFileWorkTask.data
{
    internal class StudentResult
    {
        public Theme Theme { get; set; }
        public string Question { get; set; } = string.Empty;
        public int Score { get; set; }
        public Student Student { get; set; }
        public DateOnly? Date { get; set; } = null;

        public StudentResult(Student student, Theme theme, int score, string question = "", DateOnly? date = null)
        {
            Theme = theme;
            Question = question;
            Score = score;
            Date = date;
            Student = student;
        }
    }
}
