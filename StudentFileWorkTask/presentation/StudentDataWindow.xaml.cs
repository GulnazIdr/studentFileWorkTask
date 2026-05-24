using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using StudentFileWorkTask.data;

namespace StudentFileWorkTask.presentation
{
    public partial class StudentDataWindow : Window
    {
        StudentResultViewModel studentResultViewModel;
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
    }
}
