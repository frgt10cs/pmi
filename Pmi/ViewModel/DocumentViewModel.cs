using Pmi.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Pmi.ViewModel
{
    class DocumentViewModel:BaseViewModel
    {
        private Excel excel;

        public ObservableCollection<EmployeeViewModel> Employees { get; set; } = new ObservableCollection<EmployeeViewModel>();
        public ObservableCollection<string> ReportModes { get; set; } = new ObservableCollection<string>();
        private EmployeeViewModel selectedEmployee;
        public EmployeeViewModel SelectedEmployee { get { return selectedEmployee; } set { selectedEmployee = value; OnPropertyChanged("SelectedEmployee"); } }
        private string year = "";
        public string Year { get { return year; } set { year = value; OnPropertyChanged("Year"); } }
        private string selectedMode;
        public string SelectedMode { get { return selectedMode; } set { selectedMode = value; OnPropertyChanged("SelectedMode"); } }
        private RelayCommand createReport;
        public RelayCommand CreateReport { get { return createReport; } }

        public DocumentViewModel()
        {

        }

        public bool IsYear()
        {
            Regex regex = new Regex(@"[0-9]{4}/[0-9]{4}");
            return regex.Match(year).Success;
        }
    }
}
