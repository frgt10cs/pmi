using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.Model
{
    class Discipline
    {
        public string Name { get; }
        public int Flow { get; set; }
        public int Groups { get; set; }
        public double Lectures { get; set; }
        public double PracticalWork { get; set; }
        public double LaboratoryWork { get; set; }
        public double Consultations { get; set; }
        public double WorkWithAspirants { get; set; }
        public double Coursework { get; set; }
        public double Diploms { get; set; }
        public double Practice { get; set; }
        public double Tests { get; set; }
        public double Exam { get; set; }
        public double Aspirants { get; set; }
        public double Masters { get; set; }
        public double Gak { get; set; }

        public Discipline(string name)
        {
            Name = name;
        }
        public string FlowAndCountOfGroups()
        {
            return $"{Flow}({Groups})";
        }

        public double TotalForThisDiscipline()
        {
            return Lectures + PracticalWork + LaboratoryWork + Consultations +
                WorkWithAspirants + Coursework + Diploms + Practice + Tests +
                Exam + Aspirants + Masters + Gak;
        }
    }
}
