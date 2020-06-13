using System.Collections.Generic;

namespace Pmi.Model
{
    class Discipline
    {
        public string Name { get; }
        public List<string> Groups { get; set; }
        public List<string> CodeOP { get; set; } 
        public double Lectures { get; set; }
        public double PracticalWork { get; set; }
        public double LaboratoryWork { get; set; }
        public double ConsultationsByTheory { get; set; }
        public double ConsultationsByDiplom { get; set; }
        public double Aspirants { get; set; }
        public double Coursework { get; set; }
        public double Diploms { get; set; }
        public double Practice { get; set; }
        public double GEK { get; set; }
        public double Tests { get; set; }
        public double Exam { get; set; }
        public double AnotherWork { get; set; }

        public Discipline(string name)
        {
            Name = name;
            Groups = new List<string>();
            CodeOP = new List<string>();
        }
 
        public double TotalForThisDiscipline()
        {
            return Lectures + PracticalWork + LaboratoryWork + ConsultationsByTheory +
                ConsultationsByDiplom + Coursework + Diploms + Practice + Tests +
                Exam + Aspirants + GEK + AnotherWork;
        }
    }
}
