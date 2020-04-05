using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using Pmi.Model;
using Pmi.Service.Abstraction;

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
        private Visibility mainVis;
        public Visibility MainVis { get { return mainVis; } set { mainVis = value; OnPropertyChanged("MainVis"); } }

        #region Окно Загрузки
        private Visibility loadVis;
        public Visibility LoadVis { get { return loadVis; } set { loadVis = value; OnPropertyChanged("LoadVis"); } }
        private string status;
        public string Status { get { return status; } set { status = value; OnPropertyChanged("Status"); } }
        private uint progress;
        public uint Progress { get { return progress; } set { progress = value; OnPropertyChanged("Progress"); } }
        #endregion

        #region Окно Настроек
        private Visibility settingsVis;
        public Visibility SettingsVis { get { return settingsVis; } set { settingsVis = value; OnPropertyChanged("SettingsVis"); } }
        private string icon;
        public string Icon { get { return icon; } set { icon = value; OnPropertyChanged("Icon"); } }
        private RelayCommand changeWin;
        public RelayCommand ChangeWin { get { return changeWin; } }
        #endregion

        private Excel excel;

        public bool AreYear()
        {
            Regex regex = new Regex(@"[0-9]{4}/[0-9]{4}");
            return regex.Match(year).Success;
        }

        public MainViewModel(CacheService<List<EmployeeViewModel>> cacheServ, Excel cacheExcel)
        {
            var cache = cacheServ.UploadCache();        
            if(cache==null)
            {

            }
            else
            {
                foreach(var employee in cache)
                {
                    Employees.Add(employee);
                }
            }

            MainVis = Visibility.Visible;
            LoadVis = Visibility.Hidden;
            SettingsVis = Visibility.Hidden;
            Icon = "⚙";

            excel = cacheExcel;
            excel.OnProgressChanged += (s, e) => OnPropertyChanged("Progress");
            excel.OnStatusChanged += (s, e) => OnPropertyChanged("Status");

            ReportModes = new ObservableCollection<string>()
            {
                "Сформировать отчёт по преподавателю"
            };

            createReport = new RelayCommand(obj =>
            {
                Loading();
            },
            _obj => selectedEmployee != null && selectedMode != null);

            changeWin = new RelayCommand(obj =>
            {
                if (Icon == "⚙")
                {
                    MainVis = Visibility.Hidden;
                    SettingsVis = Visibility.Visible;
                    Icon = "←";
                }
                else
                {
                    MainVis = Visibility.Visible;
                    SettingsVis = Visibility.Hidden;
                    Icon = "⚙";
                }
            },
            _obj => LoadVis == Visibility.Hidden);
        }

        private async void Loading()
        {
            Status = "";
            Progress = 0;
            MainVis = Visibility.Hidden;
            LoadVis = Visibility.Visible;
            await Task.Run(() => { excel.CreateRaportInFile("Data.xlsm", excel.GetEmployee("Data.xlsm", ref status, ref progress
                , "Заботин", "Владислав", "Иванович", "Профессор", Year), ref status, ref progress); Console.WriteLine(Progress); });
            MainVis = Visibility.Visible;
            LoadVis = Visibility.Hidden;
        }
    }
}
