using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using StudentFileWorkTask.data;

namespace StudentFileWorkTask.presentation
{
    internal class StudentResultViewModel : INotifyPropertyChanged
    {
        private bool isSumAggregationChecked;
        private bool isMiddleAggregationChecked;
        private ObservableCollection<StudentResult> InitialStudentResultList { get; set; }
        private ObservableCollection<StudentResult> AggregatedStudentResultList { get; set; }
        private ObservableCollection<StudentResult> StudentResultList { get; set; }
        private ObservableCollection<Theme> ThemeList { get; set; }
        private ObservableCollection<Group> GroupList { get; set; }
        private ObservableCollection<Student> StudentList { get; set; }

        public bool IsSumAggregationChecked
        {
            get { return isSumAggregationChecked; }
            set
            {
                isSumAggregationChecked = value;
                onPropertyChanged("IsSumAggregationChecked");
                StudentResultDataList = 
            }
        }

        public ObservableCollection<StudentResult> StudentResultDataList
        {
            get { return StudentResultList; }
            private set
            {
                StudentResultList = value;
            }
        }

        public StudentResultViewModel()
        {
            ThemeList = new ObservableCollection<Theme> { new Theme("Логика"), new Theme("Математика") };
            GroupList = new ObservableCollection<Group> { new Group("9po12"), new Group("9po14") };
            StudentList = new ObservableCollection<Student> { new Student("idrisova", "gulnaz", "raisova", GroupList[0]), new Student("idrisova2", "gulnaz2", "raisova2", GroupList[1]), };
            InitialStudentResultList = new ObservableCollection<StudentResult> {
                new StudentResult(StudentList[0], ThemeList[0], 0, "what is ur name", DateOnly.MaxValue),
                new StudentResult(StudentList[0], ThemeList[1], 0, "what is ur name", DateOnly.MaxValue),
                new StudentResult(StudentList[1], ThemeList[0], 0, "what is ur name", DateOnly.MaxValue),
                new StudentResult(StudentList[1], ThemeList[1], 0, "what is ur name", DateOnly.MaxValue),
            };

            StudentResultDataList = InitialStudentResultList;
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        public void onPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
            }
        }

        private void onAggregated()
        {
            var result = InitialStudentResultList.GroupBy(result => new { result.Student, result.Theme }).Select(r => 
                new StudentResult(
                     r.Key.Student,
                     r.Key.Theme,
                     r.Sum(s => s.Score)
                )
            );

            var list  = result.ToList<StudentResult>();
        }
    }
}
