using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows;
using StudentFileWorkTask.data;
using OfficeOpenXml;
using System.IO;
using System.Text.Json;

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
            OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            currentTemplate = new Template(true, true);
            GroupList = new ObservableCollection<Group> { new("9apo12"), new("10po14"), new("9bpo15") };
            StudentList = new ObservableCollection<Student> {
                new("bidrisova", "gulnaz", "raisova", GroupList[0]),
                new("aidrisovagul", "gulnaz", "raisova", GroupList[2]),
                new("idrisova2", "gulnaz2", "raisova2", GroupList[1])
                //new("bidrisova", "gulnaz", "raisova"),
                //new("aidrisovagul", "gulnaz", "raisova"),
                //new("idrisova2", "gulnaz2", "raisova2")
            };
            ThemeList = new ObservableCollection<Theme> { new("Логика"), new("АМатематика") };
            _questionList = new ObservableCollection<Question> { new(ThemeList[0], "vwhat is ur name1"), new(ThemeList[0], "awhat is ur name4"), new(ThemeList[1], "what is ur name12"), new(ThemeList[1], "bwhat is ur name13"), };

            List<StudentResult> resultList = new List<StudentResult>
            {
                new(StudentList[0], _questionList[1], 1, DateOnly.MaxValue),
                new(StudentList[1], _questionList[0], 1, DateOnly.MaxValue),
                new(StudentList[1], _questionList[1], 0, DateOnly.MaxValue),
                new(StudentList[0], _questionList[2], 1, DateOnly.MaxValue),
                new(StudentList[1], _questionList[2], 1, DateOnly.MaxValue),
                new(StudentList[2], _questionList[0], 1, DateOnly.MaxValue),
                new(StudentList[2], _questionList[1], 1, DateOnly.MaxValue),
                new(StudentList[2], _questionList[2], 1, DateOnly.MaxValue)
                //new(StudentList[0], _questionList[1], 1),
                //new(StudentList[1], _questionList[0], 1),
                //new(StudentList[1], _questionList[1], 0),
                //new(StudentList[0], _questionList[2], 1),
                //new(StudentList[1], _questionList[2], 1),
                //new(StudentList[2], _questionList[0], 1),
                //new(StudentList[2], _questionList[1], 1),
                //new(StudentList[2], _questionList[2], 1)
            };

            SortInitialList(resultList);
            StudentResultList = _InitialStudentResultList;

            OptionList = new ObservableCollection<StudentResultFilter>();
            AddFilterIfExists("Темы", ThemeList, t => t.ThemeName);
            AddFilterIfExists("Студенты", StudentList, s => s.Surname);
            if (currentTemplate.IsGroupExists)
            {
                AddFilterIfExists("Группы", GroupList, g => g.GroupName);
            }
        }

        public void onFilter()
        {
            var isAllUnchecked = true;
            filtered = _InitialStudentResultList;

            if (!IsDefaultggregationChecked)
            {
                UpdateDetailColumnsVisibility();
                filtered = aggregated;
            }

            foreach (var filterGroup in OptionList)
            {
                var selectedOptions = filterGroup.Options
                    .Where(o => o.IsChecked == true)
                    .Select(o => o.Option)
                    .ToList();

                if (!selectedOptions.Any()) continue;
                isAllUnchecked = false;

                switch (filterGroup.FilterName)
                {
                    case "Темы":
                        filtered = filtered.Where(res => selectedOptions.Contains(res.Question.Theme.ThemeName));
                        break;
                    case "Студенты":
                        filtered = filtered.Where(res => selectedOptions.Contains(res.Student.Surname));
                        break;
                    case "Группы":
                        filtered = filtered.Where(res => selectedOptions.Contains(res.Student.Group.GroupName));
                        break;
                }
            }

            IsFiltering = !isAllUnchecked;
            if (isAllUnchecked)
            {
                filtered = new List<StudentResult>();
                StudentResultList = _InitialStudentResultList;
                OnAggregated();
            }
            else
            {
                StudentResultList = new ObservableCollection<StudentResult>(filtered);
            }
        }

        public void OnAggregated()
        {
            UpdateDetailColumnsVisibility();
            List<StudentResult> currentList = _InitialStudentResultList.ToList();

            if (IsFiltering)
            {
                if (IsDefaultggregationChecked)
                {
                    onFilter();
                }
                currentList = filtered.ToList();
            }

            if (!IsDefaultggregationChecked)
            {
                var result = currentList.GroupBy(result => new { result.Student, result.Question.Theme }).Select(r =>
                    new StudentResult(
                         r.Key.Student,
                         new Question(r.Key.Theme),
                         r.Sum(s => s.Score)
                    )
                ).ToList();

                var list = result.ToList<StudentResult>();
                if (IsSumAggregationChecked)
                {
                    aggregated = list;
                }
                else if (IsMiddleggregationChecked)
                {
                    List<StudentResult> studentMiddleResults = new();
                    int questionsAmountPerTheme = _questionList.Count();
                    foreach (var item in list)
                    {
                        studentMiddleResults.Add(new StudentResult(item.Student, item.Question, ((double)item.Score / questionsAmountPerTheme) * 100));
                    }

                    aggregated = studentMiddleResults;
                }

                StudentResultList = new ObservableCollection<StudentResult>(aggregated);
            }
            else
            {
                StudentResultList = new ObservableCollection<StudentResult>(currentList);
            }

        }

        public List<StudentResultThemeSum> GetStudentResultThemeSummary()
        {
            var index = 0;
            if (IsSumAggregationChecked) {
                List<StudentResultThemeSum> list = aggregated.GroupBy(result => new { result.Student }).Select(r =>
                    new StudentResultThemeSum(
                         r.Key.Student,
                         r.ToDictionary(v => v.Question.Theme.ThemeName, v => (int)v.Score),
                         (int)r.Sum(s => s.Score)
                    )
                ).OrderBy(r => r.Student.Surname).ToList();

                foreach (var result in list)
                {
                    result.Index = ++index;
                }

                return list;
            }
            else
            {
                return new();
            }
        }

        private void SortInitialList(List<StudentResult> resultList)
        {
            List<StudentResult> sorted = resultList.OrderBy(r => r.Question.Quest).ToList();

            if (currentTemplate.IsDataExists)
            {
                sorted = sorted.OrderBy(r => r.Date).ToList();
            }

            sorted = sorted.OrderBy(r => r.Student.Surname).ToList();

            if (currentTemplate.IsGroupExists)
            {
                sorted = sorted.OrderBy(r => r.Student.Group.GroupName).ToList();
            }

            sorted = sorted.OrderBy(r => r.Question.Theme.ThemeName).ToList();

            _InitialStudentResultList = new ObservableCollection<StudentResult>(sorted);
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

        public event PropertyChangedEventHandler? PropertyChanged;
        private void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(prop));
        }
    
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

        public void AddFiles(string[] filePaths)
        {
            int added = 0;
            foreach (var file in filePaths)
            {
                if (!_selectedFiles.Contains(file))
                {
                    _selectedFiles.Add(file);
                    FileList.Add(System.IO.Path.GetFileName(file));
                    added++;
                }
            }
            if (added > 0)
                MessageBox.Show($"Загружено файлов: {added}");
        }

        public void AddFilesFromFolder(string folderPath)
        {
            string[] extensions = { "*.xlsx", "*.xls", "*.csv" };
            var files = new List<string>();
            foreach (var ext in extensions)
            {
                files.AddRange(System.IO.Directory.GetFiles(folderPath, ext));
            }

            int added = 0;
            foreach (var file in files)
            {
                if (!_selectedFiles.Contains(file))
                {
                    _selectedFiles.Add(file);
                    FileList.Add(System.IO.Path.GetFileName(file));
                    added++;
                }
            }
            if (added > 0)
                MessageBox.Show($"Загружено файлов из папки: {added}");
        }

        public void ClearFiles()
        {
            _selectedFiles.Clear();
            FileList.Clear();
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

        public void LoadFileWithMapping(string filePath, MappingTemplate template)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                int colCount = worksheet.Dimension.Columns;

                var headers = new List<string>();
                for (int col = 1; col <= colCount; col++)
                {
                    headers.Add(worksheet.Cells[1, col].Text);
                }

                int studentIdx = headers.IndexOf(template.StudentColumn);
                int groupIdx = headers.IndexOf(template.GroupColumn);
                int dateIdx = headers.IndexOf(template.DateColumn);

                GroupList.Clear();
                StudentList.Clear();
                ThemeList.Clear();
                _questionList.Clear();

                var newResults = new List<StudentResult>();
                var groupsDict = new Dictionary<string, Group>();
                var studentsDict = new Dictionary<string, Student>();
                var themesDict = new Dictionary<string, Theme>();
                var questionsDict = new Dictionary<string, Question>();

                for (int row = 2; row <= rowCount; row++) 
                {
                    string studentName = studentIdx >= 0 ? worksheet.Cells[row, studentIdx + 1].Text : "";
                    string groupName = groupIdx >= 0 ? worksheet.Cells[row, groupIdx + 1].Text : "";
                    string dateStr = dateIdx >= 0 ? worksheet.Cells[row, dateIdx + 1].Text : "";

                    if (string.IsNullOrEmpty(studentName)) continue;

                    var parts = studentName.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                    string surname = parts.Length > 0 ? parts[0] : "";
                    string name = parts.Length > 1 ? parts[1] : "";
                    string patronymic = parts.Length > 2 ? parts[2] : "";

                    if (!groupsDict.ContainsKey(groupName))
                        groupsDict[groupName] = new Group(string.IsNullOrEmpty(groupName) ? "Без группы" : groupName);

                    string studentKey = $"{surname}|{name}|{patronymic}";
                    if (!studentsDict.ContainsKey(studentKey))
                        studentsDict[studentKey] = new Student(surname, name, patronymic, groupsDict[groupName]);

                    DateOnly? date = null;
                    if (DateTime.TryParse(dateStr, out var dt))
                        date = DateOnly.FromDateTime(dt);

                    for (int q = 0; q < template.QuestionColumns.Count; q++)
                    {
                        string questionCol = template.QuestionColumns[q];
                        string scoreCol = template.ScoreColumns.Count > q ? template.ScoreColumns[q] : "";

                        int qIdx = headers.IndexOf(questionCol);
                        int sIdx = headers.IndexOf(scoreCol);

                        if (qIdx < 0 || sIdx < 0) continue;

                        string themeName = questionCol;
                        string questionText = worksheet.Cells[row, qIdx + 1].Text;
                        double score = double.TryParse(worksheet.Cells[row, sIdx + 1].Text, out var val) ? val : 0;

                        if (!themesDict.ContainsKey(themeName))
                            themesDict[themeName] = new Theme(themeName);

                        string questionKey = $"{themeName}|{questionText}";
                        if (!questionsDict.ContainsKey(questionKey))
                            questionsDict[questionKey] = new Question(themesDict[themeName], questionText);

                        newResults.Add(new StudentResult(studentsDict[studentKey], questionsDict[questionKey], score, date));
                    }
                }

                foreach (var g in groupsDict.Values) GroupList.Add(g);
                foreach (var s in studentsDict.Values) StudentList.Add(s);
                foreach (var t in themesDict.Values) ThemeList.Add(t);
                foreach (var q in questionsDict.Values) _questionList.Add(q);

                _InitialStudentResultList = new ObservableCollection<StudentResult>(newResults);
                SortInitialList(newResults);
                StudentResultList = _InitialStudentResultList;

                OptionList.Clear();
                AddFilterIfExists("Темы", ThemeList, t => t.ThemeName);
                AddFilterIfExists("Студенты", StudentList, s => s.Surname);
                if (GroupList.Any())
                    AddFilterIfExists("Группы", GroupList, g => g.GroupName);
            }
        }
    }
}
