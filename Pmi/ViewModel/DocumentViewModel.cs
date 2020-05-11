using Pmi.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Pmi.ViewModel
{
    class DocumentViewModel:BaseViewModel
    {
        private Excel excel;
        private EmployeeViewModel selectedEmployee;
        private string year = "";
        private string selectedMode;
        private RelayCommand createReport;
        private RelayCommand openLoadingView;
        private RelayCommand closeLoadingView;

        public ObservableCollection<EmployeeViewModel> Employees { get; set; } = new ObservableCollection<EmployeeViewModel>();
        public ObservableCollection<string> ReportModes { get; set; } = new ObservableCollection<string>();
        public EmployeeViewModel SelectedEmployee { get { return selectedEmployee; } set { selectedEmployee = value; OnPropertyChanged("SelectedEmployee"); } }
        public string Year { get { return year; } set { year = value; OnPropertyChanged("Year"); } }
        public string SelectedMode { get { return selectedMode; } set { selectedMode = value; OnPropertyChanged("SelectedMode"); } }
        public RelayCommand CreateReport { get { return createReport; } }
        public RelayCommand OpenLoadingView { get { return openLoadingView; } }
        public RelayCommand CloseLoadingView { get { return closeLoadingView; } }

        public DocumentViewModel(ObservableCollection<EmployeeViewModel> cacheEmployee, Excel cacheExcel, RelayCommand open, RelayCommand close)
        {
            Employees = cacheEmployee;
            excel = cacheExcel;
            ReportModes = new ObservableCollection<string>()
            {
                "Сформировать отчёт по преподавателю",
                "Сформировать отчёт по преподавателю (отдельный файл)"
            };

            createReport = new RelayCommand(obj =>
            {
                if (selectedMode == ReportModes[0])
                {
                    ExecuteRaportInFile();
                }
                else if (selectedMode == ReportModes[1])
                {
                    ExecuteRaportSeparate();
                }
                
            },
            _obj => selectedEmployee != null && selectedMode != null && IsYear());

            openLoadingView = open;
            closeLoadingView = close;
        }

        public bool IsYear()
        {
            Regex regex = new Regex(@"[0-9]{4}/[0-9]{4}");
            return regex.Match(year).Success;
        }

        private async void ExecuteRaportSeparate()
        {
            await Task.Run(() =>
            {
                if (ConfigurationManager.AppSettings.Get("filePath") != "")
                {
                    if (File.Exists(ConfigurationManager.AppSettings.Get("filePath")))
                    {
                        OpenLoadingView.Execute(null);
                        var employee = excel.GetEmployee(ConfigurationManager.AppSettings.Get("filePath"), new Employee(selectedEmployee), Year);
                        if (employee.AutumnSemester.Disciplines.Count != 0 || employee.SpringSemester.Disciplines.Count != 0)
                        {
                            string path = ConfigurationManager.AppSettings.Get("filePath");
                            for (int i = path.Length - 1; i > 0; i--)
                            {
                                if (path[i] == '\\')
                                {
                                    path = path.Substring(0, i);
                                    break;
                                }
                            }
                            path += "\\" + SelectedEmployee.FIO + ".xlsx";
                            excel.CreateRaportSeparate(path, employee);
                        }
                        else
                        {
                            // Employee not found
                        }
                        CloseLoadingView.Execute(null);
                    }
                    else
                    {
                        // Error: file not found
                    }
                }
                else
                {
                    // Error: file path not found
                }
            });
        }

        private async void ExecuteRaportInFile()
        {
            await Task.Run(() =>
            {
                if (ConfigurationManager.AppSettings.Get("filePath") != "")
                { 
                    if (File.Exists(ConfigurationManager.AppSettings.Get("filePath")))
                    {
                        OpenLoadingView.Execute(null);
                        var employee = excel.GetEmployee(ConfigurationManager.AppSettings.Get("filePath"), new Employee(selectedEmployee), Year);
                        if (employee.AutumnSemester.Disciplines.Count != 0 || employee.SpringSemester.Disciplines.Count != 0)
                        {
                            excel.CreateRaportInFile(ConfigurationManager.AppSettings.Get("filePath"), employee);
                        }
                        else
                        {
                            // Employee not found
                        }
                        CloseLoadingView.Execute(null);
                    }
                    else
                    {
                        // Error: file not found
                    }
                }
                else
                {
                    // Error: file path not found
                }
            });
        }
    }
}
