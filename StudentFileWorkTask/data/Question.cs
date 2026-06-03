using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentFileWorkTask.data
{
    internal class Question
    {
        public Theme Theme { get; set; }
        public string Quest { get; set; } = string.Empty;

        public Question(Theme theme, string question = "")
        {
            Quest = question;
            Theme = theme;
        }

        public override bool Equals(object? obj)
        {
            if (obj is Question question)
            {
                return Theme.ThemeName == question.Theme.ThemeName && Quest == question.Quest;
            }
            else
            {
                return false;
            }
        }
    }
}
