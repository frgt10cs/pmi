using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Pmi.ViewModel
{
    class LoadingViewModel : BaseViewModel
    {
        public event EventHandler OnRequestClose;
        public Excel ExcelObj { get; set; }

        public LoadingViewModel(Excel excel)
        {
            ExcelObj = excel;
            ExcelObj.OnProgressChanged += (s, e) => OnPropertyChanged("Status");
            ExcelObj.OnProgressChanged += (s, e) => OnPropertyChanged("Progress");
        }

        private uint progress;
        public uint Progress
        {
            get { return progress; }
            set
            {
                if (progress != value)
                {
                    progress = value;
                    OnPropertyChanged("Progress");
                }
            }
        }
        private string status;
        public string Status
        {
            get { return status; }
            set { if (status != value) { status = value; OnPropertyChanged("Status"); } }
        }
        
        public async void DoWork()
        {
            Status = "";
            Progress = 0;
            await Task.Run(() => ExcelObj.ForTest(ref status, ref progress));
            OnRequestClose(this, new EventArgs());
        }
    }
}
