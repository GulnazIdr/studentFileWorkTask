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
        //private void PdfExportBtn_Click(object sender, RoutedEventArgs e)
        //{
        //    var dialog = new SaveFileDialog
        //    {
        //        Filter = "PDF files|*.pdf",
        //        FileName = "Отчёт_студентов.pdf"
        //    };

        //    if (dialog.ShowDialog() == true)
        //    {
        //        using (var doc = new Document(PageSize.A4.Rotate()))
        //        {
        //            PdfWriter.GetInstance(doc, new FileStream(dialog.FileName, FileMode.Create));
        //            doc.Open();

        //            var font = FontFactory.GetFont(FontFactory.HELVETICA, 10);

        //            var title = new Paragraph("Отчёт по результатам тестирования");
        //            title.Alignment = Element.ALIGN_CENTER;
        //            doc.Add(title);
        //            doc.Add(new Paragraph("\n"));

        //            var table = new PdfPTable(5);
        //            table.WidthPercentage = 100;

        //            int rowNum = 1;
        //            foreach (var item in studentResultViewModel.StudentResultList)
        //            {
        //                table.AddCell(new PdfPCell(new Phrase(rowNum.ToString(), font)));
        //                table.AddCell(new PdfPCell(new Phrase(item.Student.Surname, font)));
        //                table.AddCell(new PdfPCell(new Phrase(item.Student.Group?.GroupName ?? "", font)));
        //                table.AddCell(new PdfPCell(new Phrase(item.Question.Theme?.ThemeName ?? "", font)));
        //                table.AddCell(new PdfPCell(new Phrase(item.Score.ToString(), font)));
        //                rowNum++;
        //            }

        //            doc.Add(table);
        //        }

        //        MessageBox.Show("PDF сохранён!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
        //    }
        //}

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
                    Font headerFont = new Font(baseFont, 10, Font.BOLD,BaseColor.WHITE);
                    Font cellFont = new Font(baseFont, 9, Font.NORMAL);

                    Paragraph title = new Paragraph("Отчёт по результатам тестирования", titleFont);
                    title.Alignment = Element.ALIGN_CENTER;
                    doc.Add(title);
                    doc.Add(new Paragraph("\n"));

                    PdfPTable table = new PdfPTable(6);
                    table.WidthPercentage = 100;
                    table.SetWidths(new float[] { 4f, 10f, 12f, 30f, 30f, 8f });

                    string[] headers = { "№", "Группа", "Студент", "Тема", "Вопрос", "Балл" };
                    foreach (string header in headers)
                    {
                        PdfPCell cell = new PdfPCell(new Phrase(header, headerFont));
                        cell.BackgroundColor = new BaseColor(129, 166, 198);
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.Padding = 5;
                        table.AddCell(cell);
                    }

                    int rowNum = 1;
                    foreach (var item in items)
                    {
                        PdfPCell cellNum = new PdfPCell(new Phrase(rowNum.ToString(), cellFont));
                        cellNum.Padding = 5;
                        table.AddCell(cellNum);

                        PdfPCell cellGroup = new PdfPCell(new Phrase(item.Student?.Group?.GroupName ?? "—", cellFont));
                        cellGroup.Padding = 5;
                        table.AddCell(cellGroup);

                        PdfPCell cellStudent = new PdfPCell(new Phrase(item.Student?.Surname ?? "—", cellFont));
                        cellStudent.Padding = 5;
                        table.AddCell(cellStudent);

                        PdfPCell cellTheme = new PdfPCell(new Phrase(item.Question?.Theme?.ThemeName ?? "—", cellFont));
                        cellTheme.Padding = 5;
                        cellTheme.SetLeading(0f, 1.2f);
                        table.AddCell(cellTheme);

                        PdfPCell cellQuestion = new PdfPCell(new Phrase(item.Question?.Quest ?? "—", cellFont));
                        cellQuestion.Padding = 5;
                        cellQuestion.SetLeading(0f, 1.2f);
                        table.AddCell(cellQuestion);

                        PdfPCell cellScore = new PdfPCell(new Phrase(item.Score.ToString(), cellFont));
                        cellScore.Padding = 5;
                        cellScore.HorizontalAlignment = Element.ALIGN_CENTER;
                        table.AddCell(cellScore);

                        rowNum++;
                    }

                    doc.Add(table);
                    doc.Close();

                    MessageBox.Show($"PDF успешно создан!",
                        "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
