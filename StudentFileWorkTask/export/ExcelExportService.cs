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

        public void GenerateReport(string filePath)
        {
            using (var workbook = new XLWorkbook())
            {
                var filteredData = _viewModel.StudentResultList.ToList();

                if (_viewModel.IsDefaultggregationChecked)
                {
                    CreateAllGroupsSheetDetailed(workbook, filteredData);
                    CreateGroupSheetsDetailed(workbook, filteredData);
                }
                else
                {
                    CreateAllGroupsSheetAggregated(workbook, filteredData);
                    CreateGroupSheetsAggregated(workbook, filteredData);
                }

                CreateQuestionsStatisticsSheet(workbook, filteredData);

                workbook.SaveAs(filePath);
            }
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
                    Student = $"{g.Key.Student.Surname} {g.Key.Student.Name} {g.Key.Student.Patronymic}".Trim(),
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
                        Student = $"{g.Key.Student.Surname} {g.Key.Student.Name} {g.Key.Student.Patronymic}".Trim(),
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
                    Student = $"{g.Key.Student.Surname} {g.Key.Student.Name} {g.Key.Student.Patronymic}".Trim(),
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
                        Student = $"{g.Key.Student.Surname} {g.Key.Student.Name} {g.Key.Student.Patronymic}".Trim(),
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
                .GroupBy(r => new { Theme = r.Question.Theme?.ThemeName ?? "Без темы", Question = r.Question.Quest ?? "Без названия" })
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
                worksheet.Cell(row, 4).Value = stat.TotalAnswered;
                worksheet.Cell(row, 5).Value = Math.Round(stat.Percentage, 1);
                worksheet.Cell(row, 5).Style.NumberFormat.Format = "0.0";
                row++;
            }

            worksheet.Columns().AdjustToContents();
        }
    }
}