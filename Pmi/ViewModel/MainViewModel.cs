using Pmi.Model;
using Pmi.Service.Abstraction;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.IO;

namespace Pmi.ViewModel
{
    class MainViewModel : BaseViewModel
    {

        private int openedViewIndex;
        private BaseViewModel currentViewModel;
        private string icon;
        private DocumentViewModel documentViewModel;
        private SettingsViewModel settingsViewModel;
        private LoadingViewModel loadingViewModel;
        private CacheService<List<EmployeeViewModel>> cacheService;

        public ObservableCollection<EmployeeViewModel> Employees { get; set; } = new ObservableCollection<EmployeeViewModel>();
        public int OpenedViewIndex { get { return openedViewIndex; } set { openedViewIndex = value; OnPropertyChanged("OpenedViewIndex");} }
        public BaseViewModel CurrentViewModel { get { return currentViewModel; } set { currentViewModel = value; OnPropertyChanged("CurrentViewModel"); } }        
        public string Icon { get { return icon; } set { icon = value; OnPropertyChanged("Icon"); } }


        public MainViewModel(CacheService<List<EmployeeViewModel>> cacheServ, Excel cacheExcel)
        {
            Icon = "⚙";
            cacheService = cacheServ;
            var cache = cacheServ.UploadCache();
            string path = File.Exists(ConfigurationManager.AppSettings.Get("pathCache")) ? File.ReadAllText(ConfigurationManager.AppSettings.Get("pathCache")) : "";
            ConfigurationManager.AppSettings.Set("filePath", path);
            
            if (cache == null)
            {

            }
            else
            {
                foreach (var employee in cache)
                {
                    Employees.Add(employee);
                }
            }
            
            documentViewModel = new DocumentViewModel(Employees, cacheExcel, new RelayCommand(obj =>
            {
                OpenedViewIndex = 2;
                CurrentViewModel = loadingViewModel;
            }), new RelayCommand(obj =>
            {
                OpenedViewIndex = 0;
                CurrentViewModel = documentViewModel;
            }));
            settingsViewModel = new SettingsViewModel(Employees);
            loadingViewModel = new LoadingViewModel(cacheExcel);
            OpenedViewIndex = 0;
            CurrentViewModel = documentViewModel;
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
                        Icon = "←";
                        OpenedViewIndex = 1;
                        CurrentViewModel = settingsViewModel;                        
                    }
                    else
                    {
                        Icon = "⚙";
                        OpenedViewIndex = 0;
                        CurrentViewModel = documentViewModel;
                        if (settingsViewModel.IsChanged)
                        {
                            File.WriteAllText(ConfigurationManager.AppSettings.Get("pathCache"), ConfigurationManager.AppSettings.Get("filePath"));
                            cacheService.Cache(new List<EmployeeViewModel>(Employees));
                            settingsViewModel.IsChanged = false;
                        }
                    }
                }));
            }
        }
    }
}
