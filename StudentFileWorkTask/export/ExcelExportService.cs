using ClosedXML.Excel;
using Microsoft.Win32;
using StudentFileWorkTask.data;
using StudentFileWorkTask.presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
                Filter = "Excel files (*.xlsx)|*.xlsx"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    UpdateReport(openFileDialog.FileName);
                    MessageBox.Show("Отчет успешно обновлен!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Ошибка при обновлении отчета: {ex.Message}");
                }
            }
        }

        public void GenerateReport(string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var filteredData = GetFilteredData();
                CreateAllGroupsSheet(workbook, filteredData);
                CreateGroupSheets(workbook, filteredData);
                CreateThemeSummarySheet(workbook, filteredData);
                CreateQuestionsStatisticsSheet(workbook, filteredData);

                workbook.SaveAs(filePath);
            }
        }

        private List<StudentResult> GetFilteredData()
        {
            return _viewModel.StudentResultList.ToList();
        }

        private void CreateAllGroupsSheet(XLWorkbook workbook, List<StudentResult> data)
        {
            var worksheet = workbook.Worksheets.Add("Все группы");

            worksheet.Cell(1, 1).Value = "Группа";
            worksheet.Cell(1, 2).Value = "Студент";
            worksheet.Cell(1, 3).Value = "Дата";
            worksheet.Cell(1, 4).Value = "Тема";
            worksheet.Cell(1, 5).Value = "Суммарный балл";

            var headerRange = worksheet.Range("A1:E1");
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            var groupedData = data
                .GroupBy(r => new { Group = r.Student.Group?.GroupName ?? "Без группы", Student = r.Student.Surname, Theme = r.Question.Theme.ThemeName, Date = r.Date })
                .Select(g => new
                {
                    Group = g.Key.Group,
                    Student = g.Key.Student,
                    Date = g.Key.Date,
                    Theme = g.Key.Theme,
                    TotalScore = g.Sum(r => r.Score)
                })
                .OrderBy(r => r.Group)
                .ThenBy(r => r.Student)
                .ThenBy(r => r.Date)
                .ThenBy(r => r.Theme)
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

        private void CreateGroupSheets(XLWorkbook workbook, List<StudentResult> data)
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

                worksheet.Cell(1, 1).Value = "Группа";
                worksheet.Cell(1, 2).Value = "Студент";
                worksheet.Cell(1, 3).Value = "Дата";
                worksheet.Cell(1, 4).Value = "Тема";
                worksheet.Cell(1, 5).Value = "Суммарный балл";

                var headerRange = worksheet.Range("A1:E1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                var groupedData = groupData
                    .GroupBy(r => new { Student = r.Student.Surname, Theme = r.Question.Theme.ThemeName, Date = r.Date })
                    .Select(g => new
                    {
                        Group = groupName,
                        Student = g.Key.Student,
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
                    worksheet.Cell(row, 1).Value = item.Group;
                    worksheet.Cell(row, 2).Value = item.Student;
                    worksheet.Cell(row, 3).Value = item.Date?.ToString("dd.MM.yyyy") ?? "";
                    worksheet.Cell(row, 4).Value = item.Theme;
                    worksheet.Cell(row, 5).Value = item.TotalScore;
                    row++;
                }

                worksheet.Columns().AdjustToContents();
            }
        }

        private void CreateThemeSummarySheet(XLWorkbook workbook, List<StudentResult> data)
        {
            var worksheet = workbook.Worksheets.Add("Сводка по темам");

            var allThemes = data
                .Select(r => r.Question.Theme.ThemeName)
                .Distinct()
                .OrderBy(t => t)
                .ToList();

            if (!allThemes.Any())
            {
                worksheet.Cell(1, 1).Value = "Нет данных для отображения";
                return;
            }

            var studentsData = data
                .GroupBy(r => r.Student)
                .Select(g => new
                {
                    Student = g.Key,
                    Scores = g.GroupBy(r => r.Question.Theme.ThemeName)
                              .ToDictionary(tg => tg.Key, tg => tg.Sum(r => r.Score))
                })
                .OrderBy(s => s.Student.Surname)
                .ToList();

            worksheet.Cell(1, 1).Value = "№";
            worksheet.Cell(1, 2).Value = "ФИО";

            int col = 3;
            foreach (var theme in allThemes)
            {
                worksheet.Cell(1, col).Value = theme;
                col++;
            }

            worksheet.Cell(1, col).Value = "Сумма баллов";

            var headerRange = worksheet.Range(worksheet.Cell(1, 1), worksheet.Cell(1, col));
            headerRange.Style.Font.Bold = true;
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            int row = 2;
            int index = 1;
            foreach (var studentData in studentsData)
            {
                worksheet.Cell(row, 1).Value = index++;
                worksheet.Cell(row, 2).Value = $"{studentData.Student.Surname} {studentData.Student.Name} {studentData.Student.Patronymic}".Trim();

                double totalScore = 0;
                col = 3;
                foreach (var theme in allThemes)
                {
                    double score = studentData.Scores.ContainsKey(theme) ? studentData.Scores[theme] : 0;
                    if (score > 0)
                    {
                        worksheet.Cell(row, col).Value = score;
                    }
                    totalScore += score;
                    col++;
                }

                worksheet.Cell(row, col).Value = totalScore;
                row++;
            }

            worksheet.Columns().AdjustToContents();
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
            headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
            headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

            var statistics = data
                .GroupBy(r => new { Theme = r.Question.Theme.ThemeName, Question = r.Question.Quest })
                .Select(g => new
                {
                    Theme = g.Key.Theme,
                    Question = string.IsNullOrEmpty(g.Key.Question) ? "Без названия" : g.Key.Question,
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
                worksheet.Cell(row, 4).Value = stat.TotalAnswered;
                worksheet.Cell(row, 5).Value = Math.Round(stat.Percentage, 1);
                worksheet.Cell(row, 5).Style.NumberFormat.Format = "0.0";
                row++;
            }

            worksheet.Columns().AdjustToContents();
        }

        private void UpdateReport(string filePath)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                var existingData = LoadExistingData(workbook);
                var newData = GetFilteredData();
                var mergedData = MergeData(existingData, newData);
                var sheetsToRemove = workbook.Worksheets.Where(w =>
                    w.Name == "Все группы" ||
                    w.Name == "Сводка по темам" ||
                    w.Name == "Статистика по вопросам" ||
                    (w.Name != "Все группы" && w.Name != "Сводка по темам" && w.Name != "Статистика по вопросам"));

                foreach (var sheet in sheetsToRemove.ToList())
                {
                    sheet.Delete();
                }

                CreateAllGroupsSheet(workbook, mergedData);
                CreateGroupSheets(workbook, mergedData);
                CreateThemeSummarySheet(workbook, mergedData);
                CreateQuestionsStatisticsSheet(workbook, mergedData);

                workbook.Save();
            }
        }

        private List<StudentResult> LoadExistingData(XLWorkbook workbook)
        {
            var existingData = new List<StudentResult>();

            var worksheet = workbook.Worksheet("Все группы");
            if (worksheet != null)
            {
                var lastRow = worksheet.LastRowUsed()?.RowNumber() ?? 1;
                for (int row = 2; row <= lastRow; row++)
                {
                    try
                    {
                        var groupName = worksheet.Cell(row, 1).GetString();
                        var studentName = worksheet.Cell(row, 2).GetString();
                        var dateStr = worksheet.Cell(row, 3).GetString();
                        var theme = worksheet.Cell(row, 4).GetString();
                        var score = worksheet.Cell(row, 5).GetDouble();
                        var group = string.IsNullOrEmpty(groupName) ? null : new Group(groupName);
                        var student = new Student(studentName, "", "", group);
                        var themeObj = new Theme(theme);
                        var question = new Question(themeObj, "");
                        DateOnly? date = null;
                        if (!string.IsNullOrEmpty(dateStr))
                        {
                            if (DateOnly.TryParse(dateStr, out DateOnly parsedDate))
                                date = parsedDate;
                        }

                        existingData.Add(new StudentResult(student, question, score, date));
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Ошибка загрузки строки {row}: {ex.Message}");
                    }
                }
            }

            return existingData;
        }

        private List<StudentResult> MergeData(List<StudentResult> existingData, List<StudentResult> newData)
        {
            var merged = new List<StudentResult>(existingData);

            foreach (var newItem in newData)
            {
                bool isDuplicate = existingData.Any(e =>
                    e.Student.Surname == newItem.Student.Surname &&
                    e.Student.Name == newItem.Student.Name &&
                    e.Date == newItem.Date &&
                    e.Question.Theme.ThemeName == newItem.Question.Theme.ThemeName &&
                    e.Question.Quest == newItem.Question.Quest);

                if (!isDuplicate)
                {
                    merged.Add(newItem);
                }
            }
            return merged;
        }
    }
}
