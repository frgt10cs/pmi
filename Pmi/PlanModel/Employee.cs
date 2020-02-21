using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.PlanModel
{
    class Employee
    {
        public string Name { get; set; }
        public double X { get; set; }
        public double D { get; set; }
        public Semester AutumnSemester { get; set; }
        public Semester SpringSemester { get; set; }
        public Employee(string name)
        {
            Name = name;
            AutumnSemester = new Semester();
            SpringSemester = new Semester();
        }
        public double LecturesForYear() => AutumnSemester.TotalForLectures() + SpringSemester.TotalForLectures();
        public double PracticalWorkForYear() => AutumnSemester.TotalForPracticalWork() + SpringSemester.TotalForPracticalWork();
        public double LaboratoryWorkForYear() =>
            AutumnSemester.TotalForLaboratoryWork() + SpringSemester.TotalForLaboratoryWork();

        public double ConsultationsForYear() =>
            AutumnSemester.TotalForConsultations() + SpringSemester.TotalForConsultations();

        public double WorkWithAspirantsForYear() =>
            AutumnSemester.TotalForWorkWithAspirants() + SpringSemester.TotalForWorkWithAspirants();

        public double CourseworkForYear() => AutumnSemester.TotalForCoursework() + SpringSemester.TotalForCoursework();
        public double DiplomsForYear() => AutumnSemester.TotalForDiploms() + SpringSemester.TotalForDiploms();
        public double PracticeForYear() => AutumnSemester.TotalForPractice() + SpringSemester.TotalForPractice();
        public double TestsForYear() => AutumnSemester.TotalForTests() + SpringSemester.TotalForTests();
        public double ExamForYear() => AutumnSemester.TotalForExam() + SpringSemester.TotalForExam();
        public double AspirantsForYear() => AutumnSemester.TotalForAspirants() + SpringSemester.TotalForAspirants();
        public double MastersForYear() => AutumnSemester.TotalForMasters() + SpringSemester.TotalForMasters();
        public double GakForYear() => AutumnSemester.TotalForGak() + SpringSemester.TotalForGak();

        public double Year()
        {
            return AutumnSemester.TotalForSemester() + SpringSemester.TotalForSemester();
        }
    }
}
