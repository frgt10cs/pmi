using Pmi.ViewModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.Model
{
    class EmployeeViewModel : BaseViewModel
    {
        private string lastName;
        private string firstName;
        private string patronymic;
        private string rank;
        private string studyRank;
        private string rate;
        private string staffing;

        public string FIO { get { return $"{LastName} {FirstName[0]}. {Patronymic[0]}."; } }
        public string LastName { get => lastName; set { lastName = value; OnPropertyChanged("LastName"); } }
        public string FirstName { get => firstName; set { firstName = value; OnPropertyChanged("FirstName"); } }
        public string Patronymic { get => patronymic; set { patronymic = value; OnPropertyChanged("Patronymic"); } }
        public string Rank { get => rank; set { rank = value; OnPropertyChanged("Rank"); } }
        public string StudyRank { get => studyRank; set { studyRank = value; OnPropertyChanged("StudyRank"); } }
        public string Rate { get => rate; set { rate = value; OnPropertyChanged("Rate"); } }
        public string Staffing { get => staffing; set { staffing = value; OnPropertyChanged("Staffing"); } }
    }
}
