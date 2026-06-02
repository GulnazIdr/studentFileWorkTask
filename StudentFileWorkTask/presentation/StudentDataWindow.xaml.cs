using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Win32;
using StudentFileWorkTask.data;
using StudentFileWorkTask.export;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using System.Xml.Linq;

namespace StudentFileWorkTask.presentation
{
    public partial class StudentDataWindow : Window
    {
        StudentResultViewModel studentResultViewModel;
        private List<string> _selectedFiles = new List<string>();
        private ExcelExportService exportService;

        public StudentDataWindow()
        {
            InitializeComponent();
            studentResultViewModel = new StudentResultViewModel();
            DataContext = studentResultViewModel;
            exportService = new ExcelExportService(studentResultViewModel);
        }

        private void filterCheck_Checked(object sender, RoutedEventArgs e)
        {
            studentResultViewModel.onFilter();
        }

        private void filterCheck_Unchecked(object sender, RoutedEventArgs e)
        {
            studentResultViewModel.onFilter();
        }

        private void AggregationOption_Checked(object sender, RoutedEventArgs e)
        {
            RadioButton selected = sender as RadioButton;

            if (selected == defaultAgregationOption)
            {
                studentResultGrid.Columns[studentResultGrid.Columns.Count - 1].Header = "Балл";
                studentResultViewModel.IsDefaultggregationChecked = true;
            }
            else if (selected == sumAgregationOption)
            {
                studentResultGrid.Columns[studentResultGrid.Columns.Count - 1].Header = "Суммарный балл";
                studentResultViewModel.IsSumAggregationChecked = true;
            }
            else if (selected == middleAgregationOption)
            {
                studentResultGrid.Columns[studentResultGrid.Columns.Count - 1].Header = "Процент";
                studentResultViewModel.IsMiddleggregationChecked = true;
            }

            studentResultViewModel.OnAggregated();
        }

        private void excelCreateBtn_Click(object sender, RoutedEventArgs e)
        {
            exportService.CreateNewReport();
        }

        private void BtnAddFiles_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new OpenFileDialog();
            dialog.Multiselect = true;
            dialog.Filter = "Excel files|*.xlsx;*.xls|CSV files|*.csv|All files|*.*";

            if (dialog.ShowDialog() == true)
            {
                foreach (var file in dialog.FileNames)
                {
                    var headers = studentResultViewModel.GetHeadersFromFile(file);
                    var mappingWindow = new ColumnMappingWindow(headers);

                    if (mappingWindow.ShowDialog() == true)
                    {
                        studentResultViewModel.LoadFileWithMapping(file, mappingWindow.ResultTemplate);
                    }
                }
            }
        }

        private void BtnSelectFolder_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Title = "Выберите любой файл в нужной папке";
            dialog.FileName = "выберите файл";

            if (dialog.ShowDialog() == true)
            {
                string folderPath = System.IO.Path.GetDirectoryName(dialog.FileName);
                studentResultViewModel.AddFilesFromFolder(folderPath);
            }
        }

        private void BtnClearFiles_Click(object sender, RoutedEventArgs e)
        {
            studentResultViewModel.ClearFiles();
        }

        private void PdfExportBtn_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new SaveFileDialog
            {
                Filter = "PDF files|*.pdf",
                FileName = "Отчёт_студентов.pdf"
            };

            if (dialog.ShowDialog() == true)
            {
                try
                {
                    var items = studentResultViewModel.StudentResultList;

                    if (items == null || items.Count == 0)
                    {
                        MessageBox.Show("Нет данных для экспорта!");
                        return;
                    }

                    Document doc = new Document(PageSize.A4.Rotate());
                    PdfWriter.GetInstance(doc, new FileStream(dialog.FileName, FileMode.Create));
                    doc.Open();

                    BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\arial.ttf", BaseFont.IDENTITY_H, BaseFont.EMBEDDED);
                    Font titleFont = new Font(baseFont, 16, Font.BOLD);
                    Font headerFont = new Font(baseFont, 11, Font.BOLD);
                    Font cellFont = new Font(baseFont, 9, Font.NORMAL);

                    bool isDefaultMode = studentResultViewModel.IsDefaultggregationChecked;
                    bool isSumMode = studentResultViewModel.IsSumAggregationChecked;
                    bool isMiddleMode = studentResultViewModel.IsMiddleggregationChecked;

                    string reportTitle = "";
                    if (isDefaultMode)
                        reportTitle = "Детальные результаты тестирования";
                    else if (isSumMode)
                        reportTitle = "Суммарные баллы по студентам и темам";
                    else if (isMiddleMode)
                        reportTitle = "Средние баллы (в процентах) по студентам и темам";

                    Paragraph mainTitle = new Paragraph(reportTitle, titleFont);
                    mainTitle.Alignment = Element.ALIGN_CENTER;
                    doc.Add(mainTitle);
                    doc.Add(new Paragraph("\n"));

                    if (isDefaultMode)
                    {
                        PdfPTable table = new PdfPTable(6);
                        table.WidthPercentage = 100;
                        table.SetWidths(new float[] { 4f, 10f, 12f, 30f, 30f, 8f });

                        string[] headers = { "№", "Группа", "Студент", "Тема", "Вопрос", "Балл" };
                        foreach (string header in headers)
                        {
                            Phrase phrase = new Phrase(header, headerFont);
                            PdfPCell cell = new PdfPCell(phrase);
                            cell.BackgroundColor = new BaseColor(129, 166, 198);
                            phrase.Font.Color = BaseColor.WHITE;
                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            cell.Padding = 5;
                            table.AddCell(cell);
                        }

                        int rowNum = 1;
                        foreach (var item in items)
                        {
                            table.AddCell(new PdfPCell(new Phrase(rowNum.ToString(), cellFont)) { Padding = 5 });
                            table.AddCell(new PdfPCell(new Phrase(item.Student?.Group?.GroupName ?? "—", cellFont)) { Padding = 5 });
                            table.AddCell(new PdfPCell(new Phrase(item.Student?.Surname ?? "—", cellFont)) { Padding = 5 });
                            table.AddCell(new PdfPCell(new Phrase(item.Question?.Theme?.ThemeName ?? "—", cellFont)) { Padding = 5 });
                            table.AddCell(new PdfPCell(new Phrase(item.Question?.Quest ?? "—", cellFont)) { Padding = 5 });
                            table.AddCell(new PdfPCell(new Phrase(item.Score.ToString(), cellFont)) { Padding = 5, HorizontalAlignment = Element.ALIGN_CENTER });
                            rowNum++;
                        }

                        doc.Add(table);
                    }
                    else if (isSumMode)
                    {
                        var sumData = items
                            .GroupBy(x => new { x.Student, Theme = x.Question?.Theme?.ThemeName ?? "Без темы" })
                            .Select(g => new
                            {
                                Student = g.Key.Student,
                                Theme = g.Key.Theme,
                                TotalScore = g.Sum(x => x.Score)
                            })
                            .OrderBy(x => x.Student?.Surname)
                            .ThenBy(x => x.Theme)
                            .ToList();

                        var allThemes = sumData.Select(x => x.Theme).Distinct().ToList();

                        PdfPTable sumTable = new PdfPTable(allThemes.Count + 2);
                        sumTable.WidthPercentage = 100;

                        Phrase studentPhrase = new Phrase("Студент", headerFont);
                        PdfPCell studentHeader = new PdfPCell(studentPhrase);
                        studentHeader.BackgroundColor = new BaseColor(129, 166, 198);
                        studentPhrase.Font.Color = BaseColor.WHITE;
                        studentHeader.Padding = 5;
                        sumTable.AddCell(studentHeader);

                        foreach (var theme in allThemes)
                        {
                            Phrase themePhrase = new Phrase(theme, headerFont);
                            PdfPCell themeCell = new PdfPCell(themePhrase);
                            themeCell.BackgroundColor = new BaseColor(129, 166, 198);
                            themePhrase.Font.Color = BaseColor.WHITE;
                            themeCell.Padding = 5;
                            themeCell.HorizontalAlignment = Element.ALIGN_CENTER;
                            sumTable.AddCell(themeCell);
                        }

                        Phrase totalPhrase = new Phrase("Итого", headerFont);
                        PdfPCell totalHeader = new PdfPCell(totalPhrase);
                        totalHeader.BackgroundColor = new BaseColor(129, 166, 198);
                        totalPhrase.Font.Color = BaseColor.WHITE;
                        totalHeader.Padding = 5;
                        totalHeader.HorizontalAlignment = Element.ALIGN_CENTER;
                        sumTable.AddCell(totalHeader);

                        var students = sumData.Select(x => x.Student).Distinct().OrderBy(x => x?.Surname).ToList();
                        foreach (var student in students)
                        {
                            sumTable.AddCell(new PdfPCell(new Phrase(student?.Surname ?? "—", cellFont)) { Padding = 5 });

                            double studentTotal = 0;
                            foreach (var theme in allThemes)
                            {
                                var score = sumData.FirstOrDefault(x => x.Student == student && x.Theme == theme)?.TotalScore ?? 0;
                                studentTotal += score;
                                PdfPCell scoreCell = new PdfPCell(new Phrase(score.ToString(), cellFont));
                                scoreCell.Padding = 5;
                                scoreCell.HorizontalAlignment = Element.ALIGN_CENTER;
                                sumTable.AddCell(scoreCell);
                            }

                            PdfPCell totalCell = new PdfPCell(new Phrase(studentTotal.ToString(), cellFont));
                            totalCell.Padding = 5;
                            totalCell.HorizontalAlignment = Element.ALIGN_CENTER;
                            totalCell.BackgroundColor = new BaseColor(230, 240, 250);
                            sumTable.AddCell(totalCell);
                        }

                        doc.Add(sumTable);
                    }
                    else if (isMiddleMode)
                    {
                        var maxScoresByTheme = items
                            .Where(x => x.Question?.Theme?.ThemeName != null)
                            .GroupBy(x => x.Question.Theme.ThemeName)
                            .ToDictionary(
                                g => g.Key,
                                g => g.Max(x => x.Score)
                            );

                        var avgData = items
                            .GroupBy(x => new { x.Student, Theme = x.Question?.Theme?.ThemeName ?? "Без темы" })
                            .Select(g => new
                            {
                                Student = g.Key.Student,
                                Theme = g.Key.Theme,
                                AvgPercent = maxScoresByTheme.ContainsKey(g.Key.Theme)
                                    ? g.Average(x => (x.Score / maxScoresByTheme[g.Key.Theme]) * 100)
                                    : 0
                            })
                            .OrderBy(x => x.Student?.Surname)
                            .ThenBy(x => x.Theme)
                            .ToList();

                        var allThemes = avgData.Select(x => x.Theme).Distinct().ToList();

                        PdfPTable avgTable = new PdfPTable(allThemes.Count + 2);
                        avgTable.WidthPercentage = 100;

                        Phrase avgStudentPhrase = new Phrase("Студент", headerFont);
                        PdfPCell avgStudentHeader = new PdfPCell(avgStudentPhrase);
                        avgStudentHeader.BackgroundColor = new BaseColor(129, 166, 198);
                        avgStudentPhrase.Font.Color = BaseColor.WHITE;
                        avgStudentHeader.Padding = 5;
                        avgTable.AddCell(avgStudentHeader);

                        foreach (var theme in allThemes)
                        {
                            Phrase themePhrase = new Phrase(theme, headerFont);
                            PdfPCell themeCell = new PdfPCell(themePhrase);
                            themeCell.BackgroundColor = new BaseColor(129, 166, 198);
                            themePhrase.Font.Color = BaseColor.WHITE;
                            themeCell.Padding = 5;
                            themeCell.HorizontalAlignment = Element.ALIGN_CENTER;
                            avgTable.AddCell(themeCell);
                        }

                        Phrase avgTotalPhrase = new Phrase("Средний %", headerFont);
                        PdfPCell avgTotalHeader = new PdfPCell(avgTotalPhrase);
                        avgTotalHeader.BackgroundColor = new BaseColor(129, 166, 198);
                        avgTotalPhrase.Font.Color = BaseColor.WHITE;
                        avgTotalHeader.Padding = 5;
                        avgTotalHeader.HorizontalAlignment = Element.ALIGN_CENTER;
                        avgTable.AddCell(avgTotalHeader);

                        var students = avgData.Select(x => x.Student).Distinct().OrderBy(x => x?.Surname).ToList();
                        foreach (var student in students)
                        {
                            avgTable.AddCell(new PdfPCell(new Phrase(student?.Surname ?? "—", cellFont)) { Padding = 5 });

                            double studentAvg = 0;
                            int themeCount = 0;
                            foreach (var theme in allThemes)
                            {
                                var percent = avgData.FirstOrDefault(x => x.Student == student && x.Theme == theme)?.AvgPercent ?? 0;
                                if (percent > 0) themeCount++;
                                studentAvg += percent;

                                PdfPCell percentCell = new PdfPCell(new Phrase($"{percent:F1}%", cellFont));
                                percentCell.Padding = 5;
                                percentCell.HorizontalAlignment = Element.ALIGN_CENTER;

                                if (percent >= 80)
                                    percentCell.BackgroundColor = new BaseColor(200, 230, 200);
                                else if (percent >= 60)
                                    percentCell.BackgroundColor = new BaseColor(255, 255, 200);
                                else if (percent > 0)
                                    percentCell.BackgroundColor = new BaseColor(255, 200, 200);

                                avgTable.AddCell(percentCell);
                            }

                            double finalAvg = themeCount > 0 ? studentAvg / themeCount : 0;
                            PdfPCell avgFinalCell = new PdfPCell(new Phrase($"{finalAvg:F1}%", cellFont));
                            avgFinalCell.Padding = 5;
                            avgFinalCell.HorizontalAlignment = Element.ALIGN_CENTER;
                            avgFinalCell.BackgroundColor = new BaseColor(230, 240, 250);
                            avgTable.AddCell(avgFinalCell);
                        }

                        doc.Add(avgTable);
                    }

                    doc.Close();

                    MessageBox.Show("PDF успешно создан!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
