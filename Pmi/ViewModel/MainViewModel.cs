using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Pmi.Model;

namespace Pmi.ViewModel
{
    class MainViewModel:BaseViewModel
    {
        public ObservableCollection<string> Teachers { get; set; }
        public ObservableCollection<string> ReportModes { get; set; }
        private string selectedTeacher;
        public string SelectedTeacher { get { return selectedTeacher; } set { selectedTeacher = value; OnPropertyChanged("SelectedTeacher"); } }
        private string selectedMode;
        public string SelectedMode { get { return selectedMode; } set { selectedMode = value; OnPropertyChanged("SelectedMode"); } }
        private RelayCommand createReport;
        public RelayCommand CreateReport { get { return createReport; } }

        public MainViewModel()
        {
            Teachers = new ObservableCollection<string>()
            {
                "22balin",
                "beardvedka",
                "poroshenko"
            };

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
