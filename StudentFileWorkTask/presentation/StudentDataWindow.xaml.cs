using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Win32;
using StudentFileWorkTask.data;
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

        public StudentDataWindow()
        {
            InitializeComponent();
            studentResultViewModel = new StudentResultViewModel();
            DataContext = studentResultViewModel;
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
            var dataGrid = new DataGrid();
            dataGrid.AutoGenerateColumns = false;
            dataGrid.Margin = new Thickness(10);
            dataGrid.Height = 400;

            List<Theme> themeList = studentResultViewModel.ThemeList.ToList();
            dataGrid.Columns.Add(new DataGridTextColumn()
            {
                Header = "№",
                Binding = new Binding("Index")
            });
            dataGrid.Columns.Add(new DataGridTextColumn()
            {
                Header = "ФИО",
                Binding = new Binding("Student.Surname")
            });

            for (var i = 0; i < themeList.Count; i++)
            {
                var theme = themeList[i].ThemeName;
                dataGrid.Columns.Add(new DataGridTextColumn()
                {
                    Header = theme,
                    Binding = new Binding($"[{theme}]")
                });
            }

            dataGrid.Columns.Add(new DataGridTextColumn()
            {
                Header = "Сумма баллов",
                Binding = new Binding("ScoreSummary")
            });

            List<StudentResultThemeSum> studentResultThemeSumList = studentResultViewModel.GetStudentResultThemeSummary();
            dataGrid.ItemsSource = studentResultThemeSumList;

            dataPanel.Children.Add(dataGrid);
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
                using (var doc = new Document(PageSize.A4.Rotate()))
                {
                    PdfWriter.GetInstance(doc, new FileStream(dialog.FileName, FileMode.Create));
                    doc.Open();

                    var font = FontFactory.GetFont(FontFactory.HELVETICA, 10);

                    var title = new Paragraph("Отчёт по результатам тестирования");
                    title.Alignment = Element.ALIGN_CENTER;
                    doc.Add(title);
                    doc.Add(new Paragraph("\n"));

                    var table = new PdfPTable(5);
                    table.WidthPercentage = 100;

                    int rowNum = 1;
                    foreach (var item in studentResultViewModel.StudentResultList)
                    {
                        table.AddCell(new PdfPCell(new Phrase(rowNum.ToString(), font)));
                        table.AddCell(new PdfPCell(new Phrase(item.Student.Surname, font)));
                        table.AddCell(new PdfPCell(new Phrase(item.Student.Group?.GroupName ?? "", font)));
                        table.AddCell(new PdfPCell(new Phrase(item.Question.Theme?.ThemeName ?? "", font)));
                        table.AddCell(new PdfPCell(new Phrase(item.Score.ToString(), font)));
                        rowNum++;
                    }

                    doc.Add(table);
                }

                MessageBox.Show("PDF сохранён!", "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }
}