using Pmi.Model;
using Pmi.Service.Implimentation;
using Pmi.ViewModel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Pmi
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {      
        public MainWindow()
        {
            InitializeComponent();
            Excel excel = new Excel(new JsonCacheService<StylesheetInfo>(ConfigurationManager.AppSettings.Get("stylesheetInfoCache")));
            excel.CreateRaportInFile("Test.xlsx", null);
            DataContext = new MainViewModel(new JsonCacheService<EmployeeViewModel>(ConfigurationManager.AppSettings.Get("teachersCache")));                        
        }
    }
}
