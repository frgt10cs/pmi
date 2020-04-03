using Pmi.ViewModel;
using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;

namespace Pmi
{
    /// <summary>
    /// Логика взаимодействия для LoadingWindow.xaml
    /// </summary>
    public partial class LoadingWindow : Window
    {
        public LoadingWindow()
        {
            InitializeComponent();
            Closed += (o, args) => BindableDialogResult = DialogResult;
            SetBinding(BindableDialogResultProperty, new Binding("Answer"));
        }
        void OnYes(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
        
        public bool? BindableDialogResult
        {
            get { return (bool?)GetValue(BindableDialogResultProperty); }
            set { SetValue(BindableDialogResultProperty, value); }
        }

        public static readonly DependencyProperty BindableDialogResultProperty =
            DependencyProperty.Register("BindableDialogResult", typeof(bool?), typeof(LoadingWindow),
                new FrameworkPropertyMetadata(null, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault));
    }
}
