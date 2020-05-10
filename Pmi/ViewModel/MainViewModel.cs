using Pmi.Model;
using Pmi.Service.Abstraction;
using System.Collections.Generic;

namespace Pmi.ViewModel
{
    class MainViewModel : BaseViewModel
    {

        private int openedViewIndex;
        public int OpenedViewIndex { get { return openedViewIndex; } set { openedViewIndex = value; OnPropertyChanged("OpenedViewIndex");} }

        private BaseViewModel currentViewModel;
        public BaseViewModel CurrentViewModel { get { return currentViewModel; } set { currentViewModel = value; OnPropertyChanged("CurrentViewModel"); } }        

        private string icon;
        public string Icon { get { return icon; } set { icon = value; OnPropertyChanged("Icon"); } }

        private DocumentViewModel documentViewModel;
        private SettingsViewModel settingsViewModel;
        private LoadingViewModel loadingViewModel;

        public MainViewModel(CacheService<List<EmployeeViewModel>> cacheServ, Excel cacheExcel)
        {
            Icon = "⚙";
            documentViewModel = new DocumentViewModel();
            settingsViewModel = new SettingsViewModel();
            loadingViewModel = new LoadingViewModel();
            OpenedViewIndex = 0;
            CurrentViewModel = documentViewModel;

            //var cache = cacheServ.UploadCache();        
            //if(cache==null)
            //{

            //}
            //else
            //{
            //    foreach(var employee in cache)
            //    {
            //        Employees.Add(employee);
            //    }
            //}

            //MainVis = Visibility.Visible;
            //LoadVis = Visibility.Hidden;
            //SettingsVis = Visibility.Hidden;
            //Icon = "⚙";

            //excel = cacheExcel;
            //excel.OnProgressChanged += (s, e) => OnPropertyChanged("Progress");
            //excel.OnStatusChanged += (s, e) => OnPropertyChanged("Status");

            //ReportModes = new ObservableCollection<string>()
            //{
            //    "Сформировать отчёт по преподавателю"
            //};

            //createReport = new RelayCommand(obj =>
            //{
            //    Loading();
            //},
            //_obj => selectedEmployee != null && selectedMode != null);

            //changeWin = new RelayCommand(obj =>
            //{
            //    if (Icon == "⚙")
            //    {
            //        MainVis = Visibility.Hidden;
            //        SettingsVis = Visibility.Visible;
            //        Icon = "←";
            //    }
            //    else
            //    {
            //        MainVis = Visibility.Visible;
            //        SettingsVis = Visibility.Hidden;
            //        Icon = "⚙";
            //    }
            //},
            //_obj => LoadVis == Visibility.Hidden);
        }

        private RelayCommand openSettingsView;
        public RelayCommand OpenSettingsView
        {
            get
            {
                return openSettingsView ?? (openSettingsView = new RelayCommand(obj =>
                {
                    if (CurrentViewModel == documentViewModel)
                    {
                        Icon = "Блять";
                        OpenedViewIndex = 1;
                        CurrentViewModel = settingsViewModel;                        
                    }
                    else
                    {
                        Icon = "⚙";
                        OpenedViewIndex = 0;
                        CurrentViewModel = documentViewModel;                        
                    }
                }));
            }
        }


        private async void Loading()
        {
            //Status = "";
            //Progress = 0;
            //MainVis = Visibility.Hidden;
            //LoadVis = Visibility.Visible;
            //await Task.Run(() => { excel.CreateRaportInFile("Data.xlsm", excel.GetEmployee("Data.xlsm", ref status, ref progress
            //    , "Заботин", "Владислав", "Иванович", "Профессор", Year), ref status, ref progress); Console.WriteLine(Progress); });
            //MainVis = Visibility.Visible;
            //LoadVis = Visibility.Hidden;
        }
    }
}
