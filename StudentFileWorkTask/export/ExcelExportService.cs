using ClosedXML.Excel;
using Microsoft.Win32;
using StudentFileWorkTask.data;
using StudentFileWorkTask.presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace StudentFileWorkTask.export
{
    internal class ExcelExportService
    {
        private readonly StudentResultViewModel _viewModel;

        public ExcelExportService(StudentResultViewModel viewModel)
        {
            _viewModel = viewModel;
        }

        public void CreateNewReport()
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                DefaultExt = ".xlsx",
                FileName = $"Отчет_{DateTime.Now:yyyyMMdd_HHmmss}"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                GenerateReport(saveFileDialog.FileName);
                MessageBox.Show("Отчет успешно создан!");
            }
        }

        public void UpdateExistingReport()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel files (*.xlsx)|*.xlsx",
                Title = "Выберите существующий отчет для обновления"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    var existingData = LoadExistingData(openFileDialog.FileName);
                    var newData = _viewModel.StudentResultList.ToList();
                    var mergedData = MergeData(existingData, newData);
                    UpdateExistingFile(openFileDialog.FileName, mergedData);

                    MessageBox.Show("Отчет успешно обновлен!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при обновлении отчета: {ex.Message}");
                }
            }
        }

        private void UpdateExistingFile(string filePath, List<StudentResult> data)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var sheetsToRemove = workbook.Worksheets.Where(w => w.Name != "Детальные данные").ToList();
                foreach (var sheet in sheetsToRemove)
                {
                    workbook.Worksheet(sheet.Name).Delete();
                }

                if (_viewModel.IsDefaultggregationChecked)
                {
                    CreateAllGroupsSheetDetailed(workbook, data);
                    CreateGroupSheetsDetailed(workbook, data);
                }
                else
                {
                    CreateAllGroupsSheetAggregated(workbook, data);
                    CreateGroupSheetsAggregated(workbook, data);
                }

                CreateQuestionsStatisticsSheet(workbook, data);
                UpdateDetailedDataSheet(workbook, data);

                workbook.Save();
            }
        }

        private void UpdateDetailedDataSheet(XLWorkbook workbook, List<StudentResult> data)
        {
            var worksheet = workbook.Worksheet("Детальные данные");
            if (worksheet == null)
            {
                worksheet = workbook.Worksheets.Add("Детальные данные");
                worksheet.Hide();
            }
            else
            {
                worksheet.Rows(2, worksheet.LastRowUsed()?.RowNumber() ?? 2).Delete();
            }

            worksheet.Cell(1, 1).Value = "Группа";
            worksheet.Cell(1, 2).Value = "Фамилия";
            worksheet.Cell(1, 3).Value = "Имя";
            worksheet.Cell(1, 4).Value = "Отчество";
            worksheet.Cell(1, 5).Value = "Дата";
            worksheet.Cell(1, 6).Value = "Тема";
            worksheet.Cell(1, 7).Value = "Вопрос";
            worksheet.Cell(1, 8).Value = "Балл";

            int row = 2;
            foreach (var item in data)
            {
                worksheet.Cell(row, 1).Value = item.Student.Group?.GroupName ?? "Без группы";
                worksheet.Cell(row, 2).Value = item.Student.Surname;
                worksheet.Cell(row, 3).Value = item.Student.Name;
                worksheet.Cell(row, 4).Value = item.Student.Patronymic;
                worksheet.Cell(row, 5).Value = item.Date?.ToString("dd.MM.yyyy") ?? "";
                worksheet.Cell(row, 6).Value = item.Question.Theme?.ThemeName ?? "Без темы";
                worksheet.Cell(row, 7).Value = item.Question.Quest ?? "Без вопроса";
                worksheet.Cell(row, 8).Value = item.Score;
                row++;
            }

            worksheet.Columns().AdjustToContents();
        }

        public void GenerateReport(string filePath)
        {
            var filteredData = _viewModel.StudentResultList.ToList();
            GenerateReportWithData(filePath, filteredData);
        }

        private void GenerateReportWithData(string filePath, List<StudentResult> data)
        {
            using (var workbook = new XLWorkbook())
            {
                CreateDetailedDataSheet(workbook, data);

                if (_viewModel.IsDefaultggregationChecked)
                {
                    CreateAllGroupsSheetDetailed(workbook, data);
                    CreateGroupSheetsDetailed(workbook, data);
                }
                else
                {
                    CreateAllGroupsSheetAggregated(workbook, data);
                    CreateGroupSheetsAggregated(workbook, data);
                }

                CreateQuestionsStatisticsSheet(workbook, data);

                workbook.SaveAs(filePath);
            }
        }

        private void CreateDetailedDataSheet(XLWorkbook workbook, List<StudentResult> data)
        {
            var worksheet = workbook.Worksheets.Add("Детальные данные");
            worksheet.Hide();

            worksheet.Cell(1, 1).Value = "Группа";
            worksheet.Cell(1, 2).Value = "Фамилия";
            worksheet.Cell(1, 3).Value = "Имя";
            worksheet.Cell(1, 4).Value = "Отчество";
            worksheet.Cell(1, 5).Value = "Дата";
            worksheet.Cell(1, 6).Value = "Тема";
            worksheet.Cell(1, 7).Value = "Вопрос";
            worksheet.Cell(1, 8).Value = "Балл";

            int row = 2;
            foreach (var item in data)
            {
                worksheet.Cell(row, 1).Value = item.Student.Group?.GroupName ?? "Без группы";
                worksheet.Cell(row, 2).Value = item.Student.Surname;
                worksheet.Cell(row, 3).Value = item.Student.Name;
                worksheet.Cell(row, 4).Value = item.Student.Patronymic;
                worksheet.Cell(row, 5).Value = item.Date?.ToString("dd.MM.yyyy") ?? "";
                worksheet.Cell(row, 6).Value = item.Question.Theme?.ThemeName ?? "Без темы";
                worksheet.Cell(row, 7).Value = item.Question.Quest ?? "Без вопроса";
                worksheet.Cell(row, 8).Value = item.Score;
                row++;
            }

            worksheet.Columns().AdjustToContents();
        }

        private List<StudentResult> LoadExistingData(string filePath)
        {
            var results = new List<StudentResult>();
            var groupsDict = new Dictionary<string, Group>();
            var studentsDict = new Dictionary<string, Student>();
            var themesDict = new Dictionary<string, Theme>();
            var questionsDict = new Dictionary<string, Question>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet("Детальные данные");
                if (worksheet == null) return results;

                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;

                for (int row = 2; row <= lastRow; row++)
                {
                    var groupName = worksheet.Cell(row, 1).GetString();
                    var surname = worksheet.Cell(row, 2).GetString();
                    var name = worksheet.Cell(row, 3).GetString();
                    var patronymic = worksheet.Cell(row, 4).GetString();
                    var dateStr = worksheet.Cell(row, 5).GetString();
                    var themeName = worksheet.Cell(row, 6).GetString();
                    var questionText = worksheet.Cell(row, 7).GetString();
                    var score = worksheet.Cell(row, 8).GetDouble();

                    if (string.IsNullOrEmpty(surname) || string.IsNullOrEmpty(themeName)) continue;

                    if (!groupsDict.ContainsKey(groupName))
                        groupsDict[groupName] = new Group(string.IsNullOrEmpty(groupName) ? "Без группы" : groupName);

                    string studentKey = $"{surname}|{name}|{patronymic}";
                    if (!studentsDict.ContainsKey(studentKey))
                        studentsDict[studentKey] = new Student(surname, name, patronymic, groupsDict[groupName]);

                    if (!themesDict.ContainsKey(themeName))
                        themesDict[themeName] = new Theme(themeName);

                    string questionKey = $"{themeName}|{questionText}";
                    if (!questionsDict.ContainsKey(questionKey))
                        questionsDict[questionKey] = new Question(themesDict[themeName], questionText);

                    DateOnly? date = null;
                    if (!string.IsNullOrEmpty(dateStr))
                    {
                        if (DateTime.TryParse(dateStr, out var dt))
                            date = DateOnly.FromDateTime(dt);
                        else if (DateTime.TryParseExact(dateStr, new[] { "dd.MM.yyyy", "MM/dd/yyyy", "yyyy-MM-dd" },
                                 System.Globalization.CultureInfo.InvariantCulture,
                                 System.Globalization.DateTimeStyles.None, out var dt2))
                            date = DateOnly.FromDateTime(dt2);
                    }

                    results.Add(new StudentResult(studentsDict[studentKey], questionsDict[questionKey], score, date));
                }
            }

            return results;
        }

        private List<StudentResult> MergeData(List<StudentResult> existing, List<StudentResult> newData)
        {
            var merged = new List<StudentResult>(existing);

            foreach (var newItem in newData)
            {
                bool isDuplicate = existing.Any(e =>
                    e.Student.Surname == newItem.Student.Surname &&
                    e.Student.Name == newItem.Student.Name &&
                    e.Question.Theme.ThemeName == newItem.Question.Theme.ThemeName &&
                    e.Question.Quest == newItem.Question.Quest &&
                    e.Date == newItem.Date);

                if (!isDuplicate)
                {
                    merged.Add(newItem);
                }
            }

            return merged;
        }

        private void CreateAllGroupsSheetDetailed(XLWorkbook workbook, List<StudentResult> data)
        {
            var worksheet = workbook.Worksheets.Add("Все группы");

            worksheet.Cell(1, 1).Value = "Группа";
            worksheet.Cell(1, 2).Value = "Студент";
            worksheet.Cell(1, 3).Value = "Дата";
            worksheet.Cell(1, 4).Value = "Тема";
            worksheet.Cell(1, 5).Value = "Суммарный балл";

            var headerRange = worksheet.Range("A1:E1");
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Font.FontColor = XLColor.White;
            headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(129, 166, 198);
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            var groupedData = data
                .GroupBy(r => new { Group = r.Student.Group?.GroupName ?? "Без группы", Student = r.Student, Theme = r.Question.Theme?.ThemeName ?? "Без темы", Date = r.Date })
                .Select(g => new
                {
                    Group = g.Key.Group,
                    Student = $"{g.Key.Student.Surname} {g.Key.Student.Name} {g.Key.Student.Patronymic}",
                    Date = g.Key.Date,
                    Theme = g.Key.Theme,
                    TotalScore = g.Sum(r => r.Score)
                })
                .OrderBy(r => r.Theme)
                .ThenBy(r => r.Group)
                .ThenBy(r => r.Student)
                .ThenBy(r => r.Date)
                .ToList();

            int row = 2;
            foreach (var item in groupedData)
            {
                worksheet.Cell(row, 1).Value = item.Group;
                worksheet.Cell(row, 2).Value = item.Student;
                worksheet.Cell(row, 3).Value = item.Date?.ToString("dd.MM.yyyy") ?? "";
                worksheet.Cell(row, 4).Value = item.Theme;
                worksheet.Cell(row, 5).Value = item.TotalScore;
                row++;
            }

            worksheet.Columns().AdjustToContents();
        }

        private void CreateGroupSheetsDetailed(XLWorkbook workbook, List<StudentResult> data)
        {
            var groups = data
                .Where(r => r.Student.Group != null)
                .Select(r => r.Student.Group.GroupName)
                .Distinct()
                .OrderBy(g => g);

            foreach (var groupName in groups)
            {
                var groupData = data.Where(r => r.Student.Group?.GroupName == groupName).ToList();
                var sheetName = groupName.Length > 31 ? groupName.Substring(0, 31) : groupName;
                var worksheet = workbook.Worksheets.Add(sheetName);

                worksheet.Cell(1, 1).Value = "Студент";
                worksheet.Cell(1, 2).Value = "Дата";
                worksheet.Cell(1, 3).Value = "Тема";
                worksheet.Cell(1, 4).Value = "Суммарный балл";

                var headerRange = worksheet.Range("A1:D1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(129, 166, 198);
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                var groupedData = groupData
                    .GroupBy(r => new { Student = r.Student, Theme = r.Question.Theme?.ThemeName ?? "Без темы", Date = r.Date })
                    .Select(g => new
                    {
                        Student = $"{g.Key.Student.Surname} {g.Key.Student.Name} {g.Key.Student.Patronymic}",
                        Date = g.Key.Date,
                        Theme = g.Key.Theme,
                        TotalScore = g.Sum(r => r.Score)
                    })
                    .OrderBy(r => r.Student)
                    .ThenBy(r => r.Date)
                    .ThenBy(r => r.Theme)
                    .ToList();

                int row = 2;
                foreach (var item in groupedData)
                {
                    worksheet.Cell(row, 1).Value = item.Student;
                    worksheet.Cell(row, 2).Value = item.Date?.ToString("dd.MM.yyyy") ?? "";
                    worksheet.Cell(row, 3).Value = item.Theme;
                    worksheet.Cell(row, 4).Value = item.TotalScore;
                    row++;
                }

                worksheet.Columns().AdjustToContents();
            }
        }

        private void CreateAllGroupsSheetAggregated(XLWorkbook workbook, List<StudentResult> data)
        {
            var worksheet = workbook.Worksheets.Add("Все группы");

            worksheet.Cell(1, 1).Value = "Группа";
            worksheet.Cell(1, 2).Value = "Студент";
            worksheet.Cell(1, 3).Value = "Тема";
            worksheet.Cell(1, 4).Value = "Суммарный балл";

            var headerRange = worksheet.Range("A1:D1");
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Font.FontColor = XLColor.White;
            headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(129, 166, 198);
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            var groupedData = data
                .GroupBy(r => new { Group = r.Student.Group?.GroupName ?? "Без группы", Student = r.Student, Theme = r.Question.Theme?.ThemeName ?? "Без темы" })
                .Select(g => new
                {
                    Group = g.Key.Group,
                    Student = $"{g.Key.Student.Surname} {g.Key.Student.Name} {g.Key.Student.Patronymic}",
                    Theme = g.Key.Theme,
                    TotalScore = g.Sum(r => r.Score)
                })
                .OrderBy(r => r.Theme)
                .ThenBy(r => r.Group)
                .ThenBy(r => r.Student)
                .ToList();

            int row = 2;
            foreach (var item in groupedData)
            {
                worksheet.Cell(row, 1).Value = item.Group;
                worksheet.Cell(row, 2).Value = item.Student;
                worksheet.Cell(row, 3).Value = item.Theme;
                worksheet.Cell(row, 4).Value = item.TotalScore;
                row++;
            }

            worksheet.Columns().AdjustToContents();
        }

        private void CreateGroupSheetsAggregated(XLWorkbook workbook, List<StudentResult> data)
        {
            var groups = data
                .Where(r => r.Student.Group != null)
                .Select(r => r.Student.Group.GroupName)
                .Distinct()
                .OrderBy(g => g);

            foreach (var groupName in groups)
            {
                var groupData = data.Where(r => r.Student.Group?.GroupName == groupName).ToList();
                var sheetName = groupName.Length > 31 ? groupName.Substring(0, 31) : groupName;
                var worksheet = workbook.Worksheets.Add(sheetName);

                worksheet.Cell(1, 1).Value = "Студент";
                worksheet.Cell(1, 2).Value = "Тема";
                worksheet.Cell(1, 3).Value = "Суммарный балл";

                var headerRange = worksheet.Range("A1:C1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Font.FontColor = XLColor.White;
                headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(129, 166, 198);
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                var groupedData = groupData
                    .GroupBy(r => new { Student = r.Student, Theme = r.Question.Theme?.ThemeName ?? "Без темы" })
                    .Select(g => new
                    {
                        Student = $"{g.Key.Student.Surname} {g.Key.Student.Name} {g.Key.Student.Patronymic}",
                        Theme = g.Key.Theme,
                        TotalScore = g.Sum(r => r.Score)
                    })
                    .OrderBy(r => r.Student)
                    .ThenBy(r => r.Theme)
                    .ToList();

                int row = 2;
                foreach (var item in groupedData)
                {
                    worksheet.Cell(row, 1).Value = item.Student;
                    worksheet.Cell(row, 2).Value = item.Theme;
                    worksheet.Cell(row, 3).Value = item.TotalScore;
                    row++;
                }

                worksheet.Columns().AdjustToContents();
            }
        }

        private void CreateQuestionsStatisticsSheet(XLWorkbook workbook, List<StudentResult> data)
        {
            var worksheet = workbook.Worksheets.Add("Статистика по вопросам");

            worksheet.Cell(1, 1).Value = "Тема";
            worksheet.Cell(1, 2).Value = "Вопрос";
            worksheet.Cell(1, 3).Value = "Правильных ответов";
            worksheet.Cell(1, 4).Value = "Всего ответивших";
            worksheet.Cell(1, 5).Value = "% правильных";

            var headerRange = worksheet.Range("A1:E1");
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Font.FontColor = XLColor.White;
            headerRange.Style.Fill.BackgroundColor = XLColor.FromArgb(129, 166, 198);
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            var statistics = data
                .Where(r => !string.IsNullOrWhiteSpace(r.Question.Quest))
                .GroupBy(r => new { Theme = r.Question.Theme?.ThemeName ?? "Без темы", Question = r.Question.Quest })
                .Select(g => new
                {
                    Theme = g.Key.Theme,
                    Question = g.Key.Question,
                    CorrectAnswers = g.Count(r => r.Score == 1),
                    TotalAnswered = g.Count(r => r.Score == 0 || r.Score == 1),
                    Percentage = g.Any(r => r.Score == 0 || r.Score == 1) ?
                                (double)g.Count(r => r.Score == 1) / g.Count(r => r.Score == 0 || r.Score == 1) * 100 : 0
                })
                .Where(s => s.TotalAnswered > 0)
                .OrderBy(s => s.Theme)
                .ThenBy(s => s.Question)
                .ToList();

            int row = 2;
            foreach (var stat in statistics)
            {
                worksheet.Cell(row, 1).Value = stat.Theme;
                worksheet.Cell(row, 2).Value = stat.Question;
                worksheet.Cell(row, 3).Value = stat.CorrectAnswers;
                worksheet.Cell(row, 3).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(row, 4).Value = stat.TotalAnswered;
                worksheet.Cell(row, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(row, 5).Value = Math.Round(stat.Percentage, 1);
                worksheet.Cell(row, 5).Style.NumberFormat.Format = "0.0";
                worksheet.Cell(row, 5).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                row++;
            }

            if (row == 2)
            {
                worksheet.Cell(2, 1).Value = "Нет данных для отображения";
            }

            worksheet.Columns().AdjustToContents();
        }
    }
}