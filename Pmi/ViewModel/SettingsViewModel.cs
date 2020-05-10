using Pmi.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Pmi.ViewModel
{
    class SettingsViewModel:BaseViewModel
    {
        private Visibility settingsVis;
        public Visibility SettingsVis { get { return settingsVis; } set { settingsVis = value; OnPropertyChanged("SettingsVis"); } }
        private string icon;
        public string Icon { get { return icon; } set { icon = value; OnPropertyChanged("Icon"); } }
        private RelayCommand changeWin;
        public RelayCommand ChangeWin { get { return changeWin; } }
    }
}
