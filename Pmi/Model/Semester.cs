using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.Model
{
    class Semester
    {
        public List<Discipline> Disciplines { get; set; }

        public Semester()
        {
            Disciplines = new List<Discipline>();
        }

        public double TotalForLectures() => Disciplines.Sum(a => a.Lectures);
        public double TotalForPracticalWork() => Disciplines.Sum(a => a.PracticalWork);
        public double TotalForLaboratoryWork() => Disciplines.Sum(a => a.LaboratoryWork);
        public double TotalForConsultationsByTheory() => Disciplines.Sum(a => a.ConsultationsByTheory);
        public double TotalForConsultationsByDiplom() => Disciplines.Sum(a => a.ConsultationsByDiplom);
        public double TotalForCoursework() => Disciplines.Sum(a => a.Coursework);
        public double TotalForDiploms() => Disciplines.Sum(a => a.Diploms);
        public double TotalForPractice() => Disciplines.Sum(a => a.Practice);
        public double TotalForTests() => Disciplines.Sum(a => a.Tests);
        public double TotalForExam() => Disciplines.Sum(a => a.Exam);
        public double TotalForAspirants() => Disciplines.Sum(a => a.Aspirants);
        public double TotalForGEK() => Disciplines.Sum(a => a.GEK);
        public double TotalForAnotherWork() => Disciplines.Sum(a => a.AnotherWork);
        public double TotalForSemester() => Disciplines.Sum(a => a.TotalForThisDiscipline());
    }
}
