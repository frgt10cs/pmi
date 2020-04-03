﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
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

        public bool AreYear()
        {
            Regex regex = new Regex(@"[0-9]{4}/[0-9]{4}");
            return regex.Match(year).Success;
        }

        public MainViewModel(CacheService<List<EmployeeViewModel>> cacheServ)
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

            ReportModes = new ObservableCollection<string>()
            {
                "fast",
                "slowly",
                "kirpich"
            };

            createReport = new RelayCommand(obj =>
            {
                var LoadVM = new LoadingViewModel();
                var Load = new LoadingWindow();
                Load.DataContext = LoadVM;
                LoadVM.OnRequestClose += (s, e) => Load.Close();
                Load.ShowDialog();
            },
            _obj => selectedEmployee != null && selectedMode != null );
        }                   
    }
}
