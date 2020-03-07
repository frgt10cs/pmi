using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.Model
{
    class Employee
    {
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Patronymic { get; set; }
        public string Rank { get; set; }
        public Semester AutumnSemester { get; set; }
        public Semester SpringSemester { get; set; }

        public Employee()
        {            
            AutumnSemester = new Semester();
            SpringSemester = new Semester();
        }

        public double LecturesForYear() => AutumnSemester.TotalForLectures() + SpringSemester.TotalForLectures();
        public double PracticalWorkForYear() => AutumnSemester.TotalForPracticalWork() + SpringSemester.TotalForPracticalWork();
        public double LaboratoryWorkForYear() =>
            AutumnSemester.TotalForLaboratoryWork() + SpringSemester.TotalForLaboratoryWork();

        public double ConsultationsByTheoryForYear() =>
            AutumnSemester.TotalForConsultationsByTheory() + SpringSemester.TotalForConsultationsByTheory();

        public double ConsultationsByDiplomForYear() =>
            AutumnSemester.TotalForConsultationsByDiplom() + SpringSemester.TotalForConsultationsByDiplom();

        public double CourseworkForYear() => AutumnSemester.TotalForCoursework() + SpringSemester.TotalForCoursework();
        public double DiplomsForYear() => AutumnSemester.TotalForDiploms() + SpringSemester.TotalForDiploms();
        public double PracticeForYear() => AutumnSemester.TotalForPractice() + SpringSemester.TotalForPractice();
        public double TestsForYear() => AutumnSemester.TotalForTests() + SpringSemester.TotalForTests();
        public double ExamForYear() => AutumnSemester.TotalForExam() + SpringSemester.TotalForExam();
        public double AspirantsForYear() => AutumnSemester.TotalForAspirants() + SpringSemester.TotalForAspirants();
        public double GakForYear() => AutumnSemester.TotalForGEK() + SpringSemester.TotalForGEK();
        public double AnotherWorkForYear() => AutumnSemester.TotalForAnotherWork() + SpringSemester.TotalForAnotherWork();

        public double Year()
        {
            return AutumnSemester.TotalForSemester() + SpringSemester.TotalForSemester();
        }

        public static explicit operator EmployeeViewModel(Employee employee)
        {
            return new EmployeeViewModel()
            {
                LastName = employee.LastName,
                FirstName = employee.FirstName[0],                
                Patronymic = employee.Patronymic[0],
                Rank = employee.Rank
            };
        }
    }
}
