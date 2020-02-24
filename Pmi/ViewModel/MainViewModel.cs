using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using Pmi.Model;
using Pmi.Service.Interface;

namespace Pmi.ViewModel
{
    class MainViewModel:BaseViewModel
    {
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

        public bool AreYear()
        {
            Regex regex = new Regex(@"[0-9]{4}/[0-9]{4}");
            return regex.Match(year).Success;
        }

        public MainViewModel(ICacheService<EmployeeViewModel> cacheServ)
        {
            cacheServ.UploadCache();        
            if(cacheServ.IsEmpty)
            {

            }
            else
            {
                foreach(var employee in cacheServ.GetAll())
                {
                    Employees.Add(employee);
                }
            }            

            ReportModes = new ObservableCollection<string>()
            {
                "fast",
                "slowly",
                "kirpich"
            };

            createReport = new RelayCommand(obj =>
            {
                MessageBox.Show("in the future");
            },
            _obj => selectedEmployee != null && selectedMode != null && AreYear());
        }                   
    }
}
