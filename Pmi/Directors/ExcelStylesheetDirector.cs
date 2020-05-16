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
        private readonly FontDirector fontDirector;
        private readonly ExcelFontBuilder fontBuilder;
        private readonly ExcelCellFormatBuilder cellFormatBuilder;
        private readonly ExcelCellFormatDirector cellFormatDirector;
        private readonly BorderDirector borderDirector;
        private readonly ExcelBorderBuilder borderBuilder;

        public ExcelStylesheetDirector()
        {
            //stylesheetBuilder = new ExcelStylesheetBuilder();
            fontBuilder = new ExcelFontBuilder();
            fontDirector = new FontDirector() { FontBuilder = fontBuilder };
            cellFormatBuilder = new ExcelCellFormatBuilder();
            cellFormatDirector = new ExcelCellFormatDirector() { CellFormatBuilder = cellFormatBuilder };
            borderBuilder = new ExcelBorderBuilder();
            borderDirector = new BorderDirector() { BorderBuilder = borderBuilder };
        }

        public void BuildReportStylesheet()
        {
            borderDirector.BuildEmptyBorder();
            fontDirector.BuildUniversityInfoFont();
            cellFormatDirector.BuildUniveristyInfoCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.UniveristyInfo, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildTitleFont();
            borderDirector.BuildEmptyBorder();
            cellFormatDirector.BuildTitleCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Title, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildApprove();
            borderDirector.BuildEmptyBorder();
            cellFormatDirector.BuildApproveCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Approve, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildPosition();
            borderDirector.BuildEmptyBorder();
            cellFormatDirector.BuildPositionCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Position, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildDepatment();
            borderDirector.BuildEmptyBorder();
            cellFormatDirector.BuildDepartmentCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Depatment, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildManagerInfo();
            borderDirector.BuildEmptyBorder();
            cellFormatDirector.BuildManagerInfoCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.ManagerInfo, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildManagerInfoMeta();
            borderDirector.BuildEmptyBorder();
            cellFormatDirector.BuildManagerInfoMetaCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.ManagerInfoMeta, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildTotal();
            borderDirector.BuildBorder();
            cellFormatDirector.BuildTotalCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Total, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildColumnName();
            borderDirector.BuildBorder();
            cellFormatDirector.BuildColumnNameCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.ColumnName, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildDisciplineCode();
            borderDirector.BuildBorder();
            cellFormatDirector.BuildDisciplineCodeCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.DisciplineCode, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildDisciplineName();
            borderDirector.BuildBorder();
            cellFormatDirector.BuildDisciplineNameCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.DisciplineName, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildSemesterName();
            borderDirector.BuildBorder();
            cellFormatDirector.BuildSemesterNameCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.SemesterName, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildGroupPlan();
            borderDirector.BuildBorder();
            cellFormatDirector.BuildGroupPlanCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.GroupPlan, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildColumnTotal();
            borderDirector.BuildBorder();
            cellFormatDirector.BuildColumnTotalCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.ColumnTotal, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildSemesterTotalName();
            borderDirector.BuildBorder();
            cellFormatDirector.BuildSemesterTotalLableCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.SemesterTotalLabel, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildTeacherSignature();
            borderDirector.BuildEmptyBorder();
            cellFormatDirector.BuildTeacherSignatureCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.TeacherSignature, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildYearFont();
            borderDirector.BuildEmptyBorder();
            cellFormatDirector.BuildYearCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Year, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildEmployeeInfoMeta();
            borderDirector.BuildEmptyBorder();
            cellFormatDirector.BuildEmployeeInfoMetaCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.EmployeeInfoMeta, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildEmployeeInfo();
            borderDirector.BuildEmptyBorder();
            cellFormatDirector.BuildEmployeeCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.EmployeeInfo, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildColumnNumber();
            borderDirector.BuildBorder();
            cellFormatDirector.BuildColumnNumberCellFormat(GetDefaultFont(), GetDefaultBorder());
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.ColumnNumber, cellFormatBuilder.GetCellFormat());
        }

        private uint GetDefaultFont()
        {
            return stylesheetBuilder.AddFont(fontBuilder.GetFont());
        }

        private uint GetDefaultBorder()
        {
            return stylesheetBuilder.AddBorder(borderBuilder.GetBorders());
        }
    }
}
