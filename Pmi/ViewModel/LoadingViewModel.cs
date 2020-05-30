using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Pmi.ViewModel
{
    class LoadingViewModel : BaseViewModel
    {
        private string status;
        private uint progress;

        public string Status
        {
            get => status;
            set { status = value; OnPropertyChanged("Status"); }
        }

        public uint Progress
        {
            get => progress;
            set { progress = value; OnPropertyChanged("Progress"); }
        }

        public LoadingViewModel(Excel excel)
        {
            excel.OnProgressChanged += (s, e) => { Progress = Convert.ToUInt32(s); };
            excel.OnStatusChanged += (s, e) => { Status = (string)s; };
        }
    }
}