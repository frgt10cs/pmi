﻿using Pmi.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace Pmi.ViewModel
{
    class DocumentViewModel:BaseViewModel
    {
        private readonly Excel excel;
        private EmployeeViewModel selectedEmployee;
        private string year = "";
        private string selectedMode;
        private readonly Regex yearRegex = new Regex(@"[0-9]{4}/[0-9]{4}");
        private RelayCommand createReport;
        private RelayCommand openLoadingView;
        private RelayCommand closeLoadingView;


        public ObservableCollection<EmployeeViewModel> Employees { get; set; } = new ObservableCollection<EmployeeViewModel>();

        public ObservableCollection<string> ReportModes { get; set; } = new ObservableCollection<string>();

        public EmployeeViewModel SelectedEmployee
        {
            get => selectedEmployee;
            set { selectedEmployee = value; OnPropertyChanged("SelectedEmployee"); }
        }

        public string Year
        {
            get => year;
            set { year = value; OnPropertyChanged("Year"); }
        }

        public string SelectedMode
        {
            get => selectedMode;
            set { selectedMode = value; OnPropertyChanged("SelectedMode"); }
        }

        public RelayCommand CreateReport => createReport;

        public RelayCommand OpenLoadingView => openLoadingView;

        public RelayCommand CloseLoadingView => closeLoadingView;

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
            _obj => selectedEmployee != null && selectedMode != null && IsYear()
            );

            openLoadingView = open;
            closeLoadingView = close;
        }

        public bool IsYear() => yearRegex.Match(year).Success;

        private async void ExecuteRaportSeparate()
        {
            await Task.Run(() =>
            {
                var filePath = ConfigurationManager.AppSettings.Get("filePath");

                if (filePath.Length == 0)
                {
                    MessageBox.Show("Путь к файлу не найден");
                    return;
                }
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("Файл не найден");
                    return;
                }

                OpenLoadingView.Execute(null);
                try
                {
                    var employee = excel.GetEmployee(filePath, new Employee(selectedEmployee));
                    if (!employee.HasDisciplines())
                    {
                        MessageBox.Show("Преподаватель не найден");
                        return;
                    }

                    var rightSlashPos = filePath.LastIndexOf('\\');
                    if (rightSlashPos != -1)
                    {
                        filePath = filePath.Substring(0, rightSlashPos);
                    }
                    filePath += $"\\{SelectedEmployee.FIO}.xlsx";

                    excel.CreateRaportSeparate(filePath, employee);
                }
                catch
                {
                    MessageBox.Show("Проблема с доступом к файлу");
                }
                CloseLoadingView.Execute(null);
            });
        }

        private async void ExecuteRaportInFile()
        {
            await Task.Run(() =>
            {
                var filePath = ConfigurationManager.AppSettings.Get("filePath");
                if (filePath == "")
                {
                    MessageBox.Show("Путь к файлу не найден");
                    return;
                }
                if (!File.Exists(filePath))
                {
                    MessageBox.Show("Файл не найден");
                    return;
                }

                OpenLoadingView.Execute(null);
                try
                {
                    var employee = excel.GetEmployee(filePath, new Employee(selectedEmployee));
                    if (!employee.HasDisciplines())
                    {
                        MessageBox.Show("Преподаватель не найден");
                        return;
                    }
                    excel.CreateRaportInFile(ConfigurationManager.AppSettings.Get("filePath"), employee);
                }
                catch
                {
                    MessageBox.Show("Проблема с доступом к файлу");
                }
                CloseLoadingView.Execute(null);
            });
        }
    }
}
