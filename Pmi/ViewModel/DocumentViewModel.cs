using Pmi.Model;
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
        private RelayCommand createAllReport;
        private RelayCommand openLoadingView;
        private RelayCommand closeLoadingView;
        
        private string GetYear()
        {
            if (DateTime.Now.Month < 7)
            {
                return (DateTime.Now.Year - 1).ToString() + "/" + DateTime.Now.Year;
            }
            else
            {
                return DateTime.Now.Year + "/" + (DateTime.Now.Year + 1).ToString();
            }
        }

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

        public RelayCommand CreateAllReport => createAllReport;

        public RelayCommand OpenLoadingView => openLoadingView;

        public RelayCommand CloseLoadingView => closeLoadingView;

        public DocumentViewModel(ObservableCollection<EmployeeViewModel> cacheEmployee, Excel cacheExcel, RelayCommand open, RelayCommand close)
        {
            Year = GetYear();
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

            createAllReport = new RelayCommand(obj =>
            {
                if (selectedMode == ReportModes[0])
                {
                    ExecuteAllRaportInFile();
                }
                else if (selectedMode == ReportModes[1])
                {
                    ExecuteAllRaportSeparate();
                }
            },
            _obj => selectedMode != null && IsYear()
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
                        CloseLoadingView.Execute(null);
                        return;
                    }

                    var rightSlashPos = filePath.LastIndexOf('\\');
                    if (rightSlashPos != -1)
                    {
                        filePath = filePath.Substring(0, rightSlashPos);
                    }
                    filePath += $"\\{SelectedEmployee.FIO}.xlsx";

                    excel.CreateRaportSeparate(filePath, employee, year);
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
                        CloseLoadingView.Execute(null);
                        return;
                    }
                    excel.CreateRaportInFile(ConfigurationManager.AppSettings.Get("filePath"), employee, year);
                }
                catch
                {
                    MessageBox.Show("Проблема с доступом к файлу");
                }
                CloseLoadingView.Execute(null);
            });
        }


        private async void ExecuteAllRaportSeparate()
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
                var rightSlashPos = filePath.LastIndexOf('\\');
                var _filePath = "";
                if (rightSlashPos != -1)
                {
                    _filePath = filePath.Substring(0, rightSlashPos);
                }
                OpenLoadingView.Execute(null);
                try
                {
                    foreach (var employee in Employees)
                    {
                        var Employee = excel.GetEmployee(filePath, new Employee(employee));
                        if (!Employee.HasDisciplines())
                        {
                            MessageBox.Show("Преподаватель "+employee.FIO+" не найден");
                            continue;
                        }

                        excel.CreateRaportSeparate(_filePath + $"\\{employee.FIO}.xlsx", Employee, year);
                    }
                }
                catch
                {
                    MessageBox.Show("Проблема с доступом к файлу");
                }
                CloseLoadingView.Execute(null);
            });
        }

        private async void ExecuteAllRaportInFile()
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
                    foreach (var employee in Employees)
                    {
                        var Employee = excel.GetEmployee(filePath, new Employee(employee));
                        if (!Employee.HasDisciplines())
                        {
                            MessageBox.Show("Преподаватель "+employee.FIO+" не найден");
                            continue;
                        }
                        excel.CreateRaportInFile(ConfigurationManager.AppSettings.Get("filePath"), Employee, year);
                    }
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
