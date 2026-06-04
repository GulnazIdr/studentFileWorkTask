using   System.Collections.ObjectModel;
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
using StudentFileWorkTask.domain;
using TestReporter.domain.fileWork;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace StudentFileWorkTask.presentation
{
    internal class StudentResultViewModel : INotifyPropertyChanged
    {
        private SortUitls sortUitls = new SortUitls();  
        private ObservableCollection<StudentResult> _InitialStudentResultList { get; set; }
        private ObservableCollection<Question> _questionList { get; set; }
        private Template currentTemplate { get; set; }
        private bool IsFiltering = false;
        private IEnumerable<StudentResult> filtered = new List<StudentResult>();
        private List<StudentResult> aggregated = new();

        private List<string> _loadedFiles = new List<string>();
        public List<string> LoadedFiles        
        {
            get { return _loadedFiles; }
            set { _loadedFiles = value; OnPropertyChanged(nameof(LoadedFiles)); }
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

        private ObservableCollection<DateOnly> _DateList { get; set; }
        public ObservableCollection<DateOnly> DateList
        {
            get { return _DateList; }
            private set
            {
                if (_DateList != value)
                {
                    _DateList = value;
                    OnPropertyChanged(nameof(DateList));
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
            DateList = new ObservableCollection<DateOnly>();
            _questionList = new ObservableCollection<Question>();
            _InitialStudentResultList = new ObservableCollection<StudentResult>();
            StudentResultList = new ObservableCollection<StudentResult>();
            OptionList = new ObservableCollection<StudentResultFilter>();

            UpdateFilters();
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

            }
            else
            {
                filtered = result.FilteredData;
                StudentResultList = new ObservableCollection<StudentResult>(result.FilteredData);
            }
            if(!_isDefaultAggregationChecked)
                OnAggregated();
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

            DateList.Clear();
            foreach (var t in result.Dates) DateList.Add(t);

            _questionList.Clear();
            foreach (var q in result.Questions) _questionList.Add(q);

            _InitialStudentResultList = new ObservableCollection<StudentResult>(result.Results);
            StudentResultList = _InitialStudentResultList;

            UpdateFilters();
            OnAggregated();

            IsLoading = Visibility.Collapsed;
        }

        public async void AppendMultipleFilesWithMapping(string filePath, MappingTemplate template)
        {
            IsLoading = Visibility.Visible;

            var result = await Task.Run(() => ProcessLoadMultipleFiles(filePath, template));

            foreach (var g in result.Groups) GroupList.Add(g);
            foreach (var s in result.Students) StudentList.Add(s);
            foreach (var t in result.Themes) ThemeList.Add(t);
            foreach (var t in result.Dates) DateList.Add(t);
            foreach (var q in result.Questions) _questionList.Add(q);
            SortInitialListSync(result.Results);
            foreach (var res in result.Results)
            {

                _InitialStudentResultList.Add(res);
            }

            StudentResultList = _InitialStudentResultList;

            UpdateFilters();
            OnAggregated();

            IsLoading = Visibility.Collapsed;
        }

        public void ClearFiles()
        {
            _selectedFiles.Clear();
            FileList.Clear();
        }

        public void ClearAllData()
        {
            _InitialStudentResultList.Clear();
            StudentResultList.Clear();
            GroupList.Clear();
            StudentList.Clear();
            DateList.Clear();
            ThemeList.Clear();
            _questionList.Clear();
            aggregated.Clear();
            filtered = new List<StudentResult>();
            IsFiltering = false;
            _loadedFiles.Clear();

            UpdateFilters();
        }

        public void LoadSingleFile(string filePath, MappingTemplate template)
        {
            ClearAllData();
            LoadFileWithMapping(filePath, template);
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
                var datesDict = new Dictionary<DateOnly, DateOnly>();

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
                    else
                    {
                        group = groupsDict[groupKey];
                    }

                    var studentKey = (surname, name, patronymic);
                    if (!studentsDict.TryGetValue(studentKey, out var student))
                    {
                        student = new Student(surname, name, patronymic, group);
                        studentsDict[studentKey] = student;
                    }
                    else
                    {
                        student = studentsDict[studentKey];
                    }

                    DateOnly? date = null;
                    if (DateTime.TryParse(dateStr, out var dt))
                    {
                        date = DateOnly.FromDateTime(dt);
                        if (date.HasValue && !datesDict.ContainsKey(date.Value))
                        {
                            datesDict[date.Value] = date.Value;
                        }
                    }

                    string themeName = System.IO.Path.GetFileName(filePath);
                    if (!themesDict.TryGetValue(themeName, out var theme))
                    {
                        theme = new Theme(themeName);
                        themesDict[themeName] = theme;
                    }
                    else
                    {
                        theme = themesDict[themeName];
                    }

                    for (int q = 0; q < template.QuestionColumns.Count; q++)
                    {
                        string questionCol = template.QuestionColumns[q];
                        string scoreCol = template.ScoreColumns.Count > q ? template.ScoreColumns[q] : "";

                        int qIdx = headers.IndexOf(questionCol);
                        int sIdx = headers.IndexOf(scoreCol);

                        if (qIdx < 0 || sIdx < 0) continue;

                        string questionText = worksheet.Cells[1, qIdx + 1].Text;
                        double score = double.TryParse(worksheet.Cells[row, sIdx + 1].Text, out var val) ? val : 0;

                        var questionKey = (themeName, questionText);
                        if (!questionsDict.TryGetValue(questionKey, out var question))
                        {
                            question = new Question(theme, questionText);
                            questionsDict[questionKey] = question;
                        }
                        else
                        {
                            question = questionsDict[questionKey];
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
                    Dates = datesDict.Values.ToList(),
                    Results = newResults
                };
            }
        }

        private LoadFileResult ProcessLoadMultipleFiles(string filePath, MappingTemplate template)
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
                var studentsDict = new Dictionary<string, Student>(StringComparer.Ordinal);
                var themesDict = new Dictionary<string, Theme>(StringComparer.Ordinal);
                var questionsDict = new Dictionary<string, Question>(StringComparer.Ordinal);
                var datesDict = new HashSet<DateOnly>();

                var newGroups = new List<Group>();
                var newStudents = new List<Student>();
                var newThemes = new List<Theme>();
                var newQuestions = new List<Question>();
                var newDates = new List<DateOnly>();

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

                    Group group;
                    var existingGroup = GroupList.FirstOrDefault(g => g.GroupName == groupKey);
                    if (existingGroup != null)
                    {
                        group = existingGroup;
                    }
                    else if (!groupsDict.TryGetValue(groupKey, out group))
                    {
                        group = new Group(groupKey);
                        groupsDict[groupKey] = group;
                        newGroups.Add(group);
                    }
                    else
                    {
                        group = groupsDict[groupKey];
                    }

                    string studentKey = $"{surname}|{name}|{patronymic}";
                    Student student;
                    var existingStudent = StudentList.FirstOrDefault(s => $"{s.Surname}|{s.Name}|{s.Patronymic}" == studentKey);
                    if (existingStudent != null)
                    {
                        student = existingStudent;
                    }
                    else if (!studentsDict.TryGetValue(studentKey, out student))
                    {
                        student = new Student(surname, name, patronymic, group);
                        studentsDict[studentKey] = student;
                        newStudents.Add(student);
                    }
                    else
                    {
                        student = studentsDict[studentKey];
                    }

                    DateOnly? date = null;
                    if (DateTime.TryParse(dateStr, out var dt))
                    {
                        date = DateOnly.FromDateTime(dt);
                        if (date.HasValue && !DateList.Contains(date.Value) && !datesDict.Contains(date.Value))
                        {
                            datesDict.Add(date.Value);
                            newDates.Add(date.Value);
                        }
                    }

                    string themeName = System.IO.Path.GetFileNameWithoutExtension(filePath);
                    Theme theme;
                    var existingTheme = ThemeList.FirstOrDefault(t => t.ThemeName == themeName);
                    if (existingTheme != null)
                    {
                        theme = existingTheme;
                    }
                    else if (!themesDict.TryGetValue(themeName, out theme))
                    {
                        theme = new Theme(themeName);
                        themesDict[themeName] = theme;
                        newThemes.Add(theme);
                    }
                    else
                    {
                        theme = themesDict[themeName];
                    }

                    for (int q = 0; q < template.QuestionColumns.Count; q++)
                    {
                        string questionCol = template.QuestionColumns[q];
                        string scoreCol = template.ScoreColumns.Count > q ? template.ScoreColumns[q] : "";

                        int qIdx = headers.IndexOf(questionCol);
                        int sIdx = headers.IndexOf(scoreCol);

                        if (qIdx < 0 || sIdx < 0) continue;

                        string questionText = worksheet.Cells[1, qIdx + 1].Text;
                        double score = double.TryParse(worksheet.Cells[row, sIdx + 1].Text, out var val) ? val : 0;

                        string questionKey = $"{themeName}|{questionText}";
                        Question question;
                        var existingQuestion = _questionList.FirstOrDefault(q => $"{q.Theme.ThemeName}|{q.Quest}" == questionKey);
                        if (existingQuestion != null)
                        {
                            question = existingQuestion;
                        }
                        else if (!questionsDict.TryGetValue(questionKey, out question))
                        {
                            question = new Question(theme, questionText);
                            questionsDict[questionKey] = question;
                            newQuestions.Add(question);
                        }
                        else
                        {
                            question = questionsDict[questionKey];
                        }

                        newResults.Add(new StudentResult(student, question, score, date));
                    }
                }

                return new LoadFileResult
                {
                    Groups = newGroups,
                    Students = newStudents,
                    Themes = newThemes,
                    Questions = newQuestions,
                    Dates = newDates,
                    Results = newResults
                };
            }
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

                filteredData = sortUitls.QuickFilter(filteredData, selectedOptions, filterGroup.FilterName);
            }

            return new FilterResult
            {
                FilteredData = filteredData,
                IsAllUnchecked = isAllUnchecked
            };
        }

        private List<StudentResult> ProcessAggregation(AggregationSnapshot snapshot)
        {
            var questionAmount = _questionList.Count;
            if (snapshot.IsDefaultChecked)
                return snapshot.CurrentList;

            if (snapshot.IsSumChecked)
                return sortUitls.QuickAggregation(snapshot.CurrentList, true);

            if (snapshot.IsMiddleChecked)
                return sortUitls.QuickAggregation(snapshot.CurrentList);

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

        private List<StudentResult> GetCurrentListForAggregation()
        {
            if (IsFiltering)
            {
                return filtered.ToList();
            }
  
            return _InitialStudentResultList.ToList();
        }

        private void SortInitialListSync(List<StudentResult> resultList)
        {
            sortUitls.QuickSortSync(resultList, 0, resultList.Count() - 1, isThemeName: true);

            if (currentTemplate.IsGroupExists)
            {
                sortUitls.QuickSortSync(resultList, 0, resultList.Count() - 1, isGroupName: true);
            }

            sortUitls.QuickSortSync(resultList, 0, resultList.Count() - 1, isSurname: true);
            sortUitls.QuickSortSync(resultList, 0, resultList.Count() - 1, isQuestion: true);

            if (currentTemplate.IsDataExists)
            {
                sortUitls.QuickSortSync(resultList, 0, resultList.Count() - 1);
            }
        }

        private void AddFilterIfExists<T>(string categoryName, IEnumerable<T> items, Func<T, string> nameSelector)
        {
            var filterList = items
                .Select(item => new Filter(nameSelector(item), false))
                .ToList();

            sortUitls.QuckFilterSortSync(filterList , 0, filterList.Count() - 1);
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

        private void UpdateFilters()
        {
            OptionList.Clear();
            AddFilterIfExists("Темы", ThemeList, t => t.ThemeName);
            AddFilterIfExists("Студенты", StudentList, s => s.Surname);
            if (GroupList.Any())
                AddFilterIfExists("Группы", GroupList, g => g.GroupName);
            if (DateList.Any())
                AddFilterIfExists("Даты", DateList, g => g.ToString());
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
        public List<DateOnly> Dates { get; set;  }
        public List<StudentResult> Results { get; set; }
    }
}