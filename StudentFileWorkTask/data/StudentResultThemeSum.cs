using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentFileWorkTask.data
{
    internal class StudentResultThemeSum
    {
        public int Index { get; set; } = 1;
        public Student Student { get; set; }
        public Dictionary<string, int> Score { get; set; }
        public int ScoreSummary { get; set; }

        public StudentResultThemeSum(Student student, Dictionary<string, int> score, int scoreSummary, int index = 1)
        {
            Index = index;
            Score = score;
            ScoreSummary = scoreSummary;
            Student = student;
        }

        public int this[string themeName]
        {
            get { return Score.ContainsKey(themeName) ? Score[themeName] : 0; }
        }
    }
}
