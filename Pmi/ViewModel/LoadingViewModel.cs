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

        public LoadingViewModel()
        {
            var worker = new BackgroundWorker();
            worker.DoWork += DoWork;
            worker.ProgressChanged += ProgressChanged;
            worker.WorkerReportsProgress = true;
            worker.RunWorkerAsync();
        }

        private int progress;
        public int Progress
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
        
        private void DoWork(object sender, DoWorkEventArgs e)
        {
            var worker = (BackgroundWorker)sender;
            for (int i = 0; i < 30; i++)
            {
                Progress += i;
                Thread.Sleep(100);
                
            }
            worker.ReportProgress(-1);
        }
        private void ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == -1)
            {
                OnRequestClose(this, new EventArgs());
            }
        }
    }
}
