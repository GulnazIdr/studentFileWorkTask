using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using StudentFileWorkTask.data;

namespace StudentFileWorkTask.domain
{
    internal class SortUitls
    {
        public void QuickSortSync(List<StudentResult> arr, int left, int right, bool isQuestion = false, bool isSurname = false, bool isGroupName = false, bool isThemeName = false)
        {
            if (left < right)
            {
                int pivot = 0;
                if (isQuestion || isSurname || isGroupName || isThemeName)
                    pivot = StringPartitionSync(arr, left, right, isQuestion, isSurname, isGroupName, isThemeName);
                else
                    pivot = DatePartitionSync(arr, left, right);

                QuickSortSync(arr, left, pivot - 1, isQuestion, isSurname, isGroupName, isThemeName);
                QuickSortSync(arr, pivot + 1, right, isQuestion, isSurname, isGroupName, isThemeName);
            }
        }

        public int StringPartitionSync(List<StudentResult> results, int left, int right,  bool isQuestion, bool isSurname, bool isGroupName, bool isThemeName)
        {
            string pivot = "";
            if (isQuestion) pivot = results[left].Question.Quest;
            else if (isSurname) pivot = results[left].Student.Surname;
            else if (isGroupName) pivot = results[left].Student.Group.GroupName;
            else if (isThemeName) pivot = results[left].Question.Theme.ThemeName;

            int i = left + 1;
            int j = right;

            while (i <= j)
            {
                string leftWord = "";
                if (isQuestion) leftWord = results[i].Question.Quest;
                else if (isSurname) leftWord = results[i].Student.Surname;
                else if (isGroupName) leftWord = results[i].Student.Group.GroupName;
                else if (isThemeName) leftWord = results[i].Question.Theme.ThemeName;

                while (i <= right && leftWord.CompareTo(pivot) <= 0)
                {
                    i++;
                    if (i <= right)
                    {
                        if (isQuestion) leftWord = results[i].Question.Quest;
                        else if (isSurname) leftWord = results[i].Student.Surname;
                        else if (isGroupName) leftWord = results[i].Student.Group.GroupName;
                        else if (isThemeName) leftWord = results[i].Question.Theme.ThemeName;
                    }
                }

                string rightWord = "";
                if (isQuestion) rightWord = results[j].Question.Quest;
                else if (isSurname) rightWord = results[j].Student.Surname;
                else if (isGroupName) rightWord = results[j].Student.Group.GroupName;
                else if (isThemeName) rightWord = results[j].Question.Theme.ThemeName;

                while (j > left && rightWord.CompareTo(pivot) >= 0)
                {
                    j--;
                    if (j > left)
                    {
                        if (isQuestion) rightWord = results[j].Question.Quest;
                        else if (isSurname) rightWord = results[j].Student.Surname;
                        else if (isGroupName) rightWord = results[j].Student.Group.GroupName;
                        else if (isThemeName) rightWord = results[j].Question.Theme.ThemeName;
                    }
                }

                if (i < j)
                {
                    StudentResult temp = results[i];
                    results[i] = results[j];
                    results[j] = temp;
                }

            }

            StudentResult tempPivot = results[left];
            results[left] = results[j];
            results[j] = tempPivot;
            return j;
        }

        public int DatePartitionSync(List<StudentResult> results, int left, int right)
        {
            DateOnly? pivot = results[left].Date;
            int i = left + 1;
            int j = right;

            while (i <= j)
            {
                while (i <= right && results[i].Date <= pivot)
                {
                    i++;
                }

                while (j > left && results[j].Date >= pivot)
                {
                    j--;
                }

                if (i < j)
                {
                    StudentResult? temp = results[i];
                    results[i] = results[j];
                    results[j] = temp;
                }
            }
            StudentResult tempPivot = results[left];
            results[left] = results[j];
            results[j] = tempPivot;
            return j;
        }

        public List<StudentResult> QuickFilter(List<StudentResult> initial, List<string> keyWords, string filterType)
        {
            var keyWordSet = new HashSet<string>(keyWords);
            var result = new List<StudentResult>(initial.Count);

            foreach (var item in initial)
            {
                bool matches = false;

                switch (filterType)
                {
                    case "Темы":
                        matches = keyWordSet.Contains(item.Question.Theme.ThemeName);
                        break;
                    case "Студенты":
                        matches = keyWordSet.Contains(item.Student.Surname);
                        break;
                    case "Группы":
                        matches = keyWordSet.Contains(item.Student.Group.GroupName);
                        break;
                }

                if (matches)
                    result.Add(item);
            }

            return result;
        }

        public List<StudentResult> QuickAggregation(List<StudentResult> initial, int questionsAmountPerTheme, bool isSum = false)
        {
            var dictionary = new Dictionary<string, StudentResult>();

            foreach (var item in initial)
            {
                string key = $"{item.Student.Surname}_{item.Question.Theme.ThemeName}";

                if (dictionary.TryGetValue(key, out StudentResult? existing))
                {
                    if (isSum) existing.Score += item.Score;
                    else existing.Score = ((double)item.Score / questionsAmountPerTheme) * 100;
                }
                else
                {
                    if (isSum)
                    {
                        dictionary.Add(key, new StudentResult(item.Student, new Question(item.Question.Theme), item.Score, item.Date));
                    }
                    else
                    {
                        var percent = ((double)item.Score / questionsAmountPerTheme) * 100;
                        dictionary.Add(key, new StudentResult(item.Student, new Question(item.Question.Theme), percent, item.Date));
                    }
                }
            }

            return dictionary.Values.ToList();
        }
    }
}
