using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Pmi.Model;
using Pmi.Service.Interface;

namespace Pmi.ViewModel
{
    class MainViewModel:BaseViewModel
    {
        public ObservableCollection<string> Employees { get; set; } = new ObservableCollection<string>();
        public ObservableCollection<string> ReportModes { get; set; } = new ObservableCollection<string>();
        private string selectedTeacher;
        public string SelectedTeacher { get { return selectedTeacher; } set { selectedTeacher = value; OnPropertyChanged("SelectedTeacher"); } }
        private string selectedMode;
        public string SelectedMode { get { return selectedMode; } set { selectedMode = value; OnPropertyChanged("SelectedMode"); } }
        private RelayCommand createReport;
        public RelayCommand CreateReport { get { return createReport; } }

        public MainViewModel(ICacheService<EmployeeCache> cacheServ)
        {
            cacheServ.UploadCache();        
            if(cacheServ.IsEmpty)
            {

            }
            else
            {
                foreach(var employee in cacheServ.GetAll())
                {
                    Employees.Add($"{employee.LastName} {employee.FirstName}. {employee.Patronymic}. \n{employee.Rank}");
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
            _obj => selectedTeacher != null && selectedMode != null);
        }                   
    }
}
