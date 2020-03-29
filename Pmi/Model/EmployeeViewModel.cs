using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.Model
{
    class EmployeeViewModel
    {
        public string FIO { get { return $"{LastName} {FirstName}. {Patronymic}."; } }
        public string LastName { get; set; }
        public char FirstName { get; set; }        
        public char Patronymic { get; set; }
        public string Rank { get; set; }
    }
}
