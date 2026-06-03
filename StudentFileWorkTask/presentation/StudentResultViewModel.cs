using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using StudentFileWorkTask.data;
using OfficeOpenXml;
using System.IO;
using System.Text.Json;
using DocumentFormat.OpenXml.Office.SpreadSheetML.Y2023.MsForms;
using Question = StudentFileWorkTask.data.Question;
using System.Collections.Generic;
using StudentFileWorkTask.export;

namespace StudentFileWorkTask.presentation
{
    internal class StudentResultViewModel : INotifyPropertyChanged
    {
        private ObservableCollection<StudentResult> _InitialStudentResultList { get; set; }
        private ObservableCollection<Question> _questionList { get; set; }
        private Template currentTemplate { get; set; }
        private bool IsFiltering = false;
        private IEnumerable<StudentResult> filtered = new List<StudentResult>();
        private List<StudentResult> aggregated = new();

        private List<string> _selectedFiles = new List<string>();
        public List<string> SelectedFiles
        {
            get { return _selectedFiles; }
            set { _selectedFiles = value; OnPropertyChanged(nameof(SelectedFiles)); }
        }

        private ObservableCollection<string> _fileList = new ObservableCollection<string>();
        public ObservableCollection<string> FileList
        {
            get { return _fileList; }
            set { _fileList = value; OnPropertyChanged(nameof(FileList)); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }

        private ObservableCollection<StudentResultFilter> _OptionList { get; set; }

        public ObservableCollection<StudentResultFilter> OptionList
        {
            get { return _OptionList; }
            private set
            {
                if (_OptionList != value)
                {
                    _OptionList = value;
                    OnPropertyChanged(nameof(OptionList));
                }
            }
        }

        private Visibility _isLoading;
        public Visibility IsLoading
        {
            get { return _isLoading; }
            set
            {
                if (_isLoading != value)
                {
                    _isLoading = value;
                    OnPropertyChanged(nameof(IsLoading));
                }
            }
        }

        private bool _isDefaultAggregationChecked = true;
        public bool IsDefaultggregationChecked
        {
            get { return _isDefaultAggregationChecked; }
            set
            {
                if (_isDefaultAggregationChecked != value)
                {
                    IsSumAggregationChecked = false;
                    IsMiddleggregationChecked = false;
                    _isDefaultAggregationChecked = value;
                    OnPropertyChanged(nameof(IsDefaultggregationChecked));
                }
            }
        }

        private bool _isMiddleAggregationChecked;
        public bool IsMiddleggregationChecked
        {
            get { return _isMiddleAggregationChecked; }
            set
            {
                if (_isMiddleAggregationChecked != value)
                {
                    IsDefaultggregationChecked = false;
                    IsSumAggregationChecked = false;
                    _isMiddleAggregationChecked = value;
                    OnPropertyChanged(nameof(IsMiddleggregationChecked));
                }
            }
        }

        private bool _isSumAggregationChecked;
        public bool IsSumAggregationChecked
        {
            get { return _isSumAggregationChecked; }
            set
            {
                if (_isSumAggregationChecked != value)
                {
                    IsDefaultggregationChecked = false;
                    IsMiddleggregationChecked = false;
                    _isSumAggregationChecked = value;
                    OnPropertyChanged(nameof(IsSumAggregationChecked));
                }
            }
        }

        private ObservableCollection<Theme> _ThemeList { get; set; }
        public ObservableCollection<Theme> ThemeList
        {
            get { return _ThemeList; }
            private set
            {
                if (_ThemeList != value)
                {
                    _ThemeList = value;
                    OnPropertyChanged(nameof(ThemeList));
                }
            }
        }

        private ObservableCollection<Group> _GroupList { get; set; }
        public ObservableCollection<Group> GroupList
        {
            get { return _GroupList; }
            private set
            {
                if (_GroupList != value)
                {
                    _GroupList = value;
                    OnPropertyChanged(nameof(GroupList));
                }
            }
        }

        private ObservableCollection<Student> _StudentList { get; set; }
        public ObservableCollection<Student> StudentList
        {
            get { return _StudentList; }
            private set
            {
                if (_StudentList != value)
                {
                    _StudentList = value;
                    OnPropertyChanged(nameof(StudentList));
                }
            }
        }

        private ObservableCollection<StudentResult> _StudentResultList { get; set; }
        public ObservableCollection<StudentResult> StudentResultList
        {
            get { return _StudentResultList; }
            private set
            {
                if (_StudentResultList != value)
                {
                    _StudentResultList = value;
                    OnPropertyChanged(nameof(StudentResultList));
                }
            }
        }

        private Visibility _detailColumnsVisibility;
        public Visibility DetailColumnsVisibility
        {
            get { return _detailColumnsVisibility; }
            set
            {
                _detailColumnsVisibility = value;
                OnPropertyChanged(nameof(DetailColumnsVisibility));
            }
        }

        public StudentResultViewModel()
        {
            IsLoading = Visibility.Visible;
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            currentTemplate = new Template(true, true);

            GroupList = new ObservableCollection<Group>();
            StudentList = new ObservableCollection<Student>();
            ThemeList = new ObservableCollection<Theme>();
            _questionList = new ObservableCollection<Question>();
            _InitialStudentResultList = new ObservableCollection<StudentResult>();
            StudentResultList = new ObservableCollection<StudentResult>();
            OptionList = new ObservableCollection<StudentResultFilter>();

            AddFilterIfExists("Темы", ThemeList, t => t.ThemeName);
            AddFilterIfExists("Студенты", StudentList, s => s.Surname);
            if (currentTemplate.IsGroupExists)
            {
                AddFilterIfExists("Группы", GroupList, g => g.GroupName);
            }
            IsLoading = Visibility.Collapsed;
        }

        public async void onFilter()
        {
            IsLoading = Visibility.Visible;

            var snapshot = new FilterSnapshot
            {
                InitialData = _InitialStudentResultList.ToList(),
                Options = OptionList.ToList(),
                IsDefaultChecked = IsDefaultggregationChecked
            };

            var result = await Task.Run(() => ProcessFilter(snapshot));

            IsFiltering = !result.IsAllUnchecked;
            if (result.IsAllUnchecked)
            {
                StudentResultList = _InitialStudentResultList;
                OnAggregated();
            }
            else
            {
                filtered = result.FilteredData;
                StudentResultList = new ObservableCollection<StudentResult>(result.FilteredData);
            }

            IsLoading = Visibility.Collapsed;
        }

        public async void OnAggregated()
        {
            IsLoading = Visibility.Visible;

            var snapshot = new AggregationSnapshot
            {
                IsDefaultChecked = IsDefaultggregationChecked,
                IsSumChecked = IsSumAggregationChecked,
                IsMiddleChecked = IsMiddleggregationChecked,
                CurrentList = GetCurrentListForAggregation()
            };

            var result = await Task.Run(() => ProcessAggregation(snapshot));

            StudentResultList = new ObservableCollection<StudentResult>(result);
            UpdateDetailColumnsVisibility();
            IsLoading = Visibility.Collapsed;
        }

        public async void AddFiles(string[] filePaths)
        {
            IsLoading = Visibility.Visible;

            var result = await Task.Run(() => ProcessAddFiles(filePaths, _selectedFiles));

            foreach (var file in result.NewFiles)
            {
                _selectedFiles.Add(file);
                FileList.Add(System.IO.Path.GetFileName(file));
            }
            if (result.AddedCount > 0)
                MessageBox.Show($"Загружено файлов: {result.AddedCount}");

            IsLoading = Visibility.Collapsed;
        }

        public async void AddFilesFromFolder(string folderPath)
        {
            IsLoading = Visibility.Visible;

            var result = await Task.Run(() => new ExcelImportService().ProcessAddFilesFromFolder(folderPath, _selectedFiles));

            foreach (var file in result.NewFiles)
            {
                _selectedFiles.Add(file);
                FileList.Add(System.IO.Path.GetFileName(file));
            }
            if (result.AddedCount > 0)
                MessageBox.Show($"Загружено файлов из папки: {result.AddedCount}");

            IsLoading = Visibility.Collapsed;
        }

        public async void LoadFileWithMapping(string filePath, MappingTemplate template)
        {
            IsLoading = Visibility.Visible;

            var result = await Task.Run(() => ProcessLoadFile(filePath, template, currentTemplate));

            GroupList.Clear();
            foreach (var g in result.Groups) GroupList.Add(g);

            StudentList.Clear();
            foreach (var s in result.Students) StudentList.Add(s);

            ThemeList.Clear();
            foreach (var t in result.Themes) ThemeList.Add(t);

            _questionList.Clear();
            foreach (var q in result.Questions) _questionList.Add(q);

            _InitialStudentResultList = new ObservableCollection<StudentResult>(result.Results);
            StudentResultList = _InitialStudentResultList;

            OptionList.Clear();
            AddFilterIfExists("Темы", ThemeList, t => t.ThemeName);
            AddFilterIfExists("Студенты", StudentList, s => s.Surname);
            if (GroupList.Any())
                AddFilterIfExists("Группы", GroupList, g => g.GroupName);

            IsLoading = Visibility.Collapsed;
        }

        public List<string> GetHeadersFromFile(string filePath)
        {
            var headers = new List<string>();
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.Columns;
                for (int col = 1; col <= colCount; col++)
                {
                    var header = worksheet.Cells[1, col].Text;
                    headers.Add(string.IsNullOrEmpty(header) ? $"Column{col}" : header);
                }
            }
            return headers;
        }

        public void ClearFiles()
        {
            _selectedFiles.Clear();
            FileList.Clear();
        }

        private FilterResult ProcessFilter(FilterSnapshot snapshot)
        {
            var filteredData = snapshot.InitialData;
            bool isAllUnchecked = true;

            foreach (var filterGroup in snapshot.Options)
            {
                var selectedOptions = filterGroup.Options
                    .Where(o => o.IsChecked == true)
                    .Select(o => o.Option)
                    .ToList();

                if (!selectedOptions.Any()) continue;
                isAllUnchecked = false;

                filteredData = QuickFilter(filteredData, selectedOptions, filterGroup.FilterName);
            }

            return new FilterResult
            {
                FilteredData = filteredData,
                IsAllUnchecked = isAllUnchecked
            };
        }

        private List<StudentResult> ProcessAggregation(AggregationSnapshot snapshot)
        {
            if (snapshot.IsDefaultChecked)
                return snapshot.CurrentList;

            if (snapshot.IsSumChecked)
                return QuickAggregation(snapshot.CurrentList, true);

            if (snapshot.IsMiddleChecked)
                return QuickAggregation(snapshot.CurrentList);

            return snapshot.CurrentList;
        }

        private FileAddResult ProcessAddFiles(string[] filePaths, List<string> existingFiles)
        {
            var newFiles = new List<string>();
            int added = 0;

            foreach (var file in filePaths)
            {
                if (!existingFiles.Contains(file))
                {
                    newFiles.Add(file);
                    added++;
                }
            }

            return new FileAddResult { NewFiles = newFiles, AddedCount = added };
        }

        private LoadFileResult ProcessLoadFile(string filePath, MappingTemplate template, Template currentTemplate)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                var headers = new List<string>(colCount);
                for (int col = 1; col <= colCount; col++)
                {
                    headers.Add(worksheet.Cells[1, col].Text);
                }

                int studentIdx = headers.IndexOf(template.StudentColumn);
                int groupIdx = headers.IndexOf(template.GroupColumn);
                int dateIdx = headers.IndexOf(template.DateColumn);

                var groupsDict = new Dictionary<string, Group>(StringComparer.Ordinal);
                var studentsDict = new Dictionary<(string, string, string), Student>();
                var themesDict = new Dictionary<string, Theme>(StringComparer.Ordinal);
                var questionsDict = new Dictionary<(string, string), Question>();

                int estimatedRows = rowCount - 1;
                var newResults = new List<StudentResult>(estimatedRows * template.QuestionColumns.Count);

                for (int row = 2; row <= rowCount; row++)
                {
                    string studentName = studentIdx >= 0 ? worksheet.Cells[row, studentIdx + 1].Text : "";
                    if (string.IsNullOrEmpty(studentName)) continue;

                    string groupName = groupIdx >= 0 ? worksheet.Cells[row, groupIdx + 1].Text : "";
                    string dateStr = dateIdx >= 0 ? worksheet.Cells[row, dateIdx + 1].Text : "";

                    var parts = studentName.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                    string surname = parts.Length > 0 ? parts[0] : "";
                    string name = parts.Length > 1 ? parts[1] : "";
                    string patronymic = parts.Length > 2 ? parts[2] : "";

                    string groupKey = string.IsNullOrEmpty(groupName) ? "Без группы" : groupName;
                    if (!groupsDict.TryGetValue(groupKey, out var group))
                    {
                        group = new Group(groupKey);
                        groupsDict[groupKey] = group;
                    }

                    var studentKey = (surname, name, patronymic);
                    if (!studentsDict.TryGetValue(studentKey, out var student))
                    {
                        student = new Student(surname, name, patronymic, group);
                        studentsDict[studentKey] = student;
                    }

                    DateOnly? date = null;
                    if (DateTime.TryParse(dateStr, out var dt))
                        date = DateOnly.FromDateTime(dt);

                    string themeName = System.IO.Path.GetFileName(filePath);
                    if (!themesDict.TryGetValue(themeName, out var theme))
                    {
                        theme = new Theme(themeName);
                        themesDict[themeName] = theme;
                    }

                    for (int q = 0; q < template.QuestionColumns.Count; q++)
                    {
                        string questionCol = template.QuestionColumns[q];
                        string scoreCol = template.ScoreColumns.Count > q ? template.ScoreColumns[q] : "";

                        int qIdx = headers.IndexOf(questionCol);
                        int sIdx = headers.IndexOf(scoreCol);

                        if (qIdx < 0 || sIdx < 0) continue;

                        string questionText = worksheet.Cells[row, qIdx + 1].Text;
                        double score = double.TryParse(worksheet.Cells[row, sIdx + 1].Text, out var val) ? val : 0;

                        var questionKey = (themeName, questionText);
                        if (!questionsDict.TryGetValue(questionKey, out var question))
                        {
                            question = new Question(theme, questionText);
                            questionsDict[questionKey] = question;
                        }

                        newResults.Add(new StudentResult(student, question, score, date));
                    }
                }

                SortInitialListSync(newResults);

                return new LoadFileResult
                {
                    Groups = groupsDict.Values.ToList(),
                    Students = studentsDict.Values.ToList(),
                    Themes = themesDict.Values.ToList(),
                    Questions = questionsDict.Values.ToList(),
                    Results = newResults
                };
            }
        }

        private List<StudentResult> GetCurrentListForAggregation()
        {
            if (IsFiltering && IsDefaultggregationChecked)
            {
                return filtered.ToList();
            }
            return _InitialStudentResultList.ToList();
        }

        private void SortInitialListSync(List<StudentResult> resultList)
        {
            QuickSortSync(resultList, 0, resultList.Count() - 1, isThemeName: true);

            if (currentTemplate.IsGroupExists)
            {
                QuickSortSync(resultList, 0, resultList.Count() - 1, isGroupName: true);
            }

            QuickSortSync(resultList, 0, resultList.Count() - 1, isSurname: true);
            QuickSortSync(resultList, 0, resultList.Count() - 1, isQuestion: true);

            if (currentTemplate.IsDataExists)
            {
                QuickSortSync(resultList, 0, resultList.Count() - 1);
            }
        }

        private void AddFilterIfExists<T>(string categoryName, IEnumerable<T> items, Func<T, string> nameSelector)
        {
            var filterList = items
                .Select(item => new Filter(nameSelector(item), false))
                .ToList();

            if (filterList.Any() && filterList != null)
            {
                OptionList.Add(new StudentResultFilter(categoryName, filterList));
            }
        }

        private void UpdateDetailColumnsVisibility()
        {
            DetailColumnsVisibility = (!IsDefaultggregationChecked)
                ? Visibility.Collapsed
                : Visibility.Visible;
        }

        private void QuickSortSync(List<StudentResult> arr, int left, int right, bool isQuestion = false, bool isSurname = false, bool isGroupName = false, bool isThemeName = false)
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

        private int StringPartitionSync(List<StudentResult> results, int left, int right, bool isQuestion, bool isSurname, bool isGroupName, bool isThemeName)
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

                SortInitialList(_InitialStudentResultList.ToList());
                StudentResultList = _InitialStudentResultList;

                UpdateFilters();
            }

            StudentResult tempPivot = results[left];
            results[left] = results[j];
            results[j] = tempPivot;
            return j;
        }

        private int DatePartitionSync(List<StudentResult> results, int left, int right)
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

        private List<StudentResult> QuickFilter(List<StudentResult> initial, List<string> keyWords, string filterType)
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

        private List<StudentResult> QuickAggregation(List<StudentResult> initial, bool isSum = false)
        {
            int questionsAmountPerTheme = _questionList.Count();
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

    internal class FilterSnapshot
    {
        public List<StudentResult> InitialData { get; set; }
        public List<StudentResultFilter> Options { get; set; }
        public bool IsDefaultChecked { get; set; }
    }

    internal class FilterResult
    {
        public List<StudentResult> FilteredData { get; set; }
        public bool IsAllUnchecked { get; set; }
    }

    internal class AggregationSnapshot
    {
        public bool IsDefaultChecked { get; set; }
        public bool IsSumChecked { get; set; }
        public bool IsMiddleChecked { get; set; }
        public List<StudentResult> CurrentList { get; set; }
    }

    internal class FileAddResult
    {
        public List<string> NewFiles { get; set; }
        public int AddedCount { get; set; }
    }

    internal class LoadFileResult
    {
        public List<Group> Groups { get; set; }
        public List<Student> Students { get; set; }
        public List<Theme> Themes { get; set; }
        public List<Question> Questions { get; set; }
        public List<StudentResult> Results { get; set; }
    }
}