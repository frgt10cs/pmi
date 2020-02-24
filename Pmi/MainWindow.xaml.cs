using Pmi.Service.Implimentation;
using Pmi.Service.Interface;
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
        class Teacher
        {
            public string Name { get; set; }
            public string Qualification { get; set; }
        }
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainViewModel();
            var cacheServ = new JsonCacheService<Teacher>( System.IO.Path.GetFullPath( System.IO.Path.Combine(Directory.GetCurrentDirectory(), @ConfigurationManager.AppSettings.Get("teachersCache"))));
            cacheServ.UploadCache();
            Console.WriteLine(cacheServ.GetAll().Count);
            cacheServ.Add(new Teacher() { Name = "vvv", Qualification = "qqq" });
            cacheServ.SaveChanges();
            Console.WriteLine(cacheServ.GetAll().Count);
            cacheServ.UploadCache();
            Console.WriteLine(cacheServ.GetAll().Count);
        }
    }
}
