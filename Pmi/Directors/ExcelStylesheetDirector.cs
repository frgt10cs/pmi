using Pmi.Builders;
using Pmi.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public enum ExcelCellFormats
{
    UniveristyInfo = 0,
    Title = 1,
    Approve = 2,
    Position = 3,
    Depatment = 4,
    ManagerInfo = 5,
    ManagerInfoMeta = 6,
    Total = 7,
    ColumnName = 8,
    DisciplineCode = 9,
    DisciplineName = 10,
    SemesterName = 11,
    GroupPlan = 12,
    ColumnTotal = 13,
    SemesterTotalLabel = 14,
    TeacherSignature = 15,
    Year = 16,
    EmployeeInfoMeta = 17,
    EmployeeInfo = 18,
    ColumnNumber = 19
}

namespace Pmi.Directors
{    
    public class ExcelStylesheetDirector
    {
        private ExcelStylesheetBuilder stylesheetBuilder;
        public ExcelStylesheetBuilder StylesheetBuilder { set { stylesheetBuilder = value; } }
        private FontDirector fontDirector;
        private ExcelFontBuilder fontBuilder;
        private ExcelCellFormatBuilder cellFormatBuilder;
        private ExcelCellFormatDirector cellFormatDirector;

        public ExcelStylesheetDirector()
        {
            //stylesheetBuilder = new ExcelStylesheetBuilder();
            fontBuilder = new ExcelFontBuilder();
            fontDirector = new FontDirector() { FontBuilder = fontBuilder };
            cellFormatBuilder = new ExcelCellFormatBuilder();
            cellFormatDirector = new ExcelCellFormatDirector() { CellFormatBuilder = cellFormatBuilder };
        }

        public void BuildReportStylesheet()
        {
            fontDirector.BuildUniversityInfoFont();
            cellFormatDirector.BuildUniveristyInfoCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.UniveristyInfo, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildTitleFont();
            cellFormatDirector.BuildTitleCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Title, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildApprove();
            cellFormatDirector.BuildApproveCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Approve, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildPosition();
            cellFormatDirector.BuildPositionCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Position, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildDepatment();
            cellFormatDirector.BuildDepartmentCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Depatment, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildManagerInfo();
            cellFormatDirector.BuildManagerInfoCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.ManagerInfo, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildManagerInfoMeta();
            cellFormatDirector.BuildManagerInfoMetaCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.ManagerInfoMeta, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildTotal();
            cellFormatDirector.BuildTotalCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Total, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildColumnName();
            cellFormatDirector.BuildColumnNameCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.ColumnName, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildDisciplineCode();
            cellFormatDirector.BuildDisciplineCodeCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.DisciplineCode, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildDisciplineName();
            cellFormatDirector.BuildDisciplineNameCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.DisciplineName, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildSemesterName();
            cellFormatDirector.BuildSemesterNameCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.SemesterName, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildGroupPlan();
            cellFormatDirector.BuildGroupPlanCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.GroupPlan, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildColumnTotal();
            cellFormatDirector.BuildColumnTotalCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.ColumnTotal, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildSemesterTotalName();
            cellFormatDirector.BuildSemesterTotalLableCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.SemesterTotalLabel, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildTeacherSignature();
            cellFormatDirector.BuildTeacherSignatureCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.TeacherSignature, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildYearFont();
            cellFormatDirector.BuildYearCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Year, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildEmployeeInfoMeta();
            cellFormatDirector.BuildEmployeeInfoMetaCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.EmployeeInfoMeta, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildEmployeeInfo();
            cellFormatDirector.BuildEmployeeCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.EmployeeInfo, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildColumnNumber();
            cellFormatDirector.BuildColumnNumberCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.ColumnNumber, cellFormatBuilder.GetCellFormat());
        }
    }
}
