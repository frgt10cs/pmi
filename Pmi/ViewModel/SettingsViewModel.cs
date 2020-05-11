using Microsoft.Win32;
using Pmi.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Pmi.ViewModel
{
    class SettingsViewModel : BaseViewModel
    {
        public ObservableCollection<EmployeeViewModel> Employees { get; set; }
        private EmployeeViewModel selectedEmployee;
        private string icon;
        private string filePath;
        private RelayCommand review;
        private string fio;
        private string rank;
        private string studyRank;
        private string rate;
        private string staffing;
        private RelayCommand add;
        private RelayCommand remove;
        private RelayCommand change;
        public bool IsChanged { get; set; } = false;

        public EmployeeViewModel SelectedEmployee { get { return selectedEmployee; } set
            {
                selectedEmployee = value;
                OnPropertyChanged("SelectedEmployee");
                if (value != null)
                {
                    Fio = $"{selectedEmployee.LastName} {selectedEmployee.FirstName} {selectedEmployee.Patronymic}";
                    Rank = selectedEmployee.Rank;
                    StudyRank = selectedEmployee.StudyRank;
                    Rate = selectedEmployee.Rate;
                    Staffing = selectedEmployee.Staffing;
                }
            }
        }
        public string Icon { get { return icon; } set { icon = value; OnPropertyChanged("Icon"); } }
        public string FilePath { get { return filePath; } set { filePath = value; OnPropertyChanged("FilePath"); } }
        public RelayCommand Rewiew
        {
            get
            {
                return review ?? (review = new RelayCommand(obj =>
                {
                    OpenFileDialog openFileDialog = new OpenFileDialog();
                    openFileDialog.Filter = "Excel documents (*.xlsx,*.xlsm)|*.xlsx;*.xlsm";
                    if (openFileDialog.ShowDialog() == true)
                    {
                        FilePath = openFileDialog.FileName;
                        ConfigurationManager.AppSettings.Set("filePath", FilePath);
                        IsChanged = true;
                    }
                }));
            }
        }

        public string Fio { get { return fio; } set { fio = value; OnPropertyChanged("Fio"); } }
        public string Rank { get { return rank; } set { rank = value; OnPropertyChanged("Rank"); } }
        public string StudyRank { get { return studyRank; } set { studyRank = value; OnPropertyChanged("StudyRank"); } }
        public string Rate { get { return rate; } set { rate = value; OnPropertyChanged("Rate"); } }
        public string Staffing { get { return staffing; } set { staffing = value; OnPropertyChanged("Staffing"); } }
        public RelayCommand Add
        {
            get
            {
                return add ?? (add = new RelayCommand(obj =>
                {
                    var temp = fio.Split(' ');
                    if (temp.Length == 3)
                    {
                        var TempEmployee = new EmployeeViewModel()
                        {
                            LastName = temp[0],
                            FirstName = temp[1],
                            Patronymic = temp[2],
                            Rank = rank,
                            StudyRank = studyRank,
                            Rate = rate,
                            Staffing = staffing
                        };
                        Employees.Add(TempEmployee);
                        SelectedEmployee = TempEmployee;
                        IsChanged = true;
                    }
                    else
                    {
                        // invalid format of fio
                    }
                }));
            }
        }
        public RelayCommand Remove
        {
            get
            {
                return remove ?? (remove = new RelayCommand(obj =>
                {
                    if (selectedEmployee != null)
                    {
                        Employees.Remove(selectedEmployee);
                        SelectedEmployee = null;
                        IsChanged = true;
                    }
                }));
            }
        }
        public RelayCommand Change { get { return change; } }

        public SettingsViewModel(ObservableCollection<EmployeeViewModel> cacheEmployee)
        {
            Employees = cacheEmployee;
            FilePath = ConfigurationManager.AppSettings.Get("filePath");

            change = new RelayCommand(obj =>
            {
                var temp = fio.Split(' ');
                if (temp.Length == 3)
                {
                    selectedEmployee.LastName = temp[0];
                    selectedEmployee.FirstName = temp[1];
                    selectedEmployee.Patronymic = temp[2];
                    selectedEmployee.Rank = rank;
                    selectedEmployee.StudyRank = studyRank;
                    selectedEmployee.Rate = rate;
                    selectedEmployee.Staffing = staffing;
                    IsChanged = true;
                }
                else
                {
                    // invalid format of fio
                }
            },
            _obj => selectedEmployee != null);
        }
    }
}
