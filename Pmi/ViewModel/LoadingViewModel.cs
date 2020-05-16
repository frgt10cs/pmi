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
        public string Status { get { return status; } set { status = value; OnPropertyChanged("Status"); } }
        private uint progress;
        public uint Progress { get { return progress; } set { progress = value; OnPropertyChanged("Progress"); } }

        public LoadingViewModel(Excel excel)
        {
            excel.OnProgressChanged += (s, e) => { Progress = Convert.ToUInt32(s); };
            excel.OnStatusChanged += (s, e) => { Status = (string)s; };
        }
    }
}