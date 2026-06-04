using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
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
                {
                    pivot = StudentResultPartitionSync(arr, left, right, isQuestion, isSurname, isGroupName, isThemeName);

                }
                else
                    pivot = DatePartitionSync(arr, left, right);

                QuickSortSync(arr, left, pivot - 1, isQuestion, isSurname, isGroupName, isThemeName);
                QuickSortSync(arr, pivot + 1, right, isQuestion, isSurname, isGroupName, isThemeName);
            }
        }

        public void QuckFilterSortSync(List<Filter> arr, int left, int right)
        {
            if (left < right)
            {
                int pivot = FilterPartitionSync(arr, left, right);

                QuckFilterSortSync(arr, left, pivot - 1);
                QuckFilterSortSync(arr, pivot + 1, right);
            }
        }

        public int StudentResultPartitionSync(List<StudentResult> results, int left, int right, bool isQuestion, bool isSurname, bool isGroupName, bool isThemeName)
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
                while (i <= right)
                {
                    string current = "";
                    if (isQuestion) current = results[i].Question.Quest;
                    else if (isSurname) current = results[i].Student.Surname;
                    else if (isGroupName) current = results[i].Student.Group.GroupName;
                    else if (isThemeName) current = results[i].Question.Theme.ThemeName;

                    if (current.CompareTo(pivot) > 0) break; 
                    i++;
                }

                while (j > left)
                {
                    string current = "";
                    if (isQuestion) current = results[j].Question.Quest;
                    else if (isSurname) current = results[j].Student.Surname;
                    else if (isGroupName) current = results[j].Student.Group.GroupName;
                    else if (isThemeName) current = results[j].Question.Theme.ThemeName;

                    if (current.CompareTo(pivot) < 0) break; 
                    j--;
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

        public int FilterPartitionSync(List<Filter> results, int left, int right)
        {
            string pivot = results[left].Option;

            int i = left + 1;
            int j = right;

            while (i <= j)
            {
                while (i <= right)
                {
                    if (results[i].Option.CompareTo(pivot) > 0) break;
                    i++;
                }

                while (j > left)
                {

                    if (results[j].Option.CompareTo(pivot) < 0) break;
                    j--;
                }

                if (i < j)
                {
                    Filter temp = results[i];
                    results[i] = results[j];
                    results[j] = temp;
                }
            }

            Filter tempPivot = results[left];
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
                    case "Даты":
                        matches = keyWordSet.Contains(item.Date.ToString());
                        break;
                }

                if (matches)
                    result.Add(item);
            }

            return result;
        }

        public List<StudentResult> QuickAggregation(List<StudentResult> initial, bool isSum = false)
        {
          //  MessageBox.Show("here");
            var dictionary = new Dictionary<string, StudentResult>();

            foreach (var item in initial)
            {
                string key = $"{item.Student.Surname}_{item.Question.Theme.ThemeName}";
                int questionAmount = initial.Where(i => i.Question.Theme == item.Question.Theme).Count();

                if (dictionary.TryGetValue(key, out StudentResult? existing))
                {
                    if (isSum) existing.Score += item.Score;
                    else existing.Score = ((double)item.Score / questionAmount) * 100;
                }
                else
                {
                    if (isSum)
                    {
                        dictionary.Add(key, new StudentResult(item.Student, new Question(item.Question.Theme), item.Score, item.Date));
                    }
                    else
                    {
                        var percent = ((double)item.Score / questionAmount) * 100;
                        dictionary.Add(key, new StudentResult(item.Student, new Question(item.Question.Theme), percent, item.Date));
                    }
                }
            }

            return dictionary.Values.ToList();
        }
    }
}
