using Pmi.Builders;

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
        public ExcelStylesheetBuilder StylesheetBuilder { set => stylesheetBuilder = value; }
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
            fontDirector.BuildUniversityInfoFont();
            cellFormatDirector.BuildUniveristyInfoCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.UniveristyInfo, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildTitleFont();            
            cellFormatDirector.BuildTitleCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Title, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildApprove();            
            cellFormatDirector.BuildApproveCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Approve, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildPosition();            
            cellFormatDirector.BuildPositionCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Position, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildDepatment();            
            cellFormatDirector.BuildDepartmentCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Depatment, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildManagerInfo();            
            cellFormatDirector.BuildManagerInfoCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.ManagerInfo, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildManagerInfoMeta();            
            cellFormatDirector.BuildManagerInfoMetaCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()),stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.ManagerInfoMeta, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildTotal();
            borderDirector.BuildDefaultBorders();
            cellFormatDirector.BuildTotalCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Total, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildColumnName();
            borderDirector.BuildDefaultBorders();
            cellFormatDirector.BuildColumnNameCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.ColumnName, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildDisciplineCode();
            borderDirector.BuildDefaultBorders();
            cellFormatDirector.BuildDisciplineCodeCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()),stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.DisciplineCode, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildDisciplineName();
            borderDirector.BuildDefaultBorders();
            cellFormatDirector.BuildDisciplineNameCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.DisciplineName, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildSemesterName();
            borderDirector.BuildDefaultBorders();
            cellFormatDirector.BuildSemesterNameCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.SemesterName, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildGroupPlan();
            borderDirector.BuildDefaultBorders();
            cellFormatDirector.BuildGroupPlanCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.GroupPlan, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildColumnTotal();
            borderDirector.BuildDefaultBorders();
            cellFormatDirector.BuildColumnTotalCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()),stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.ColumnTotal, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildSemesterTotalName();
            borderDirector.BuildDefaultBorders();
            cellFormatDirector.BuildSemesterTotalLableCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.SemesterTotalLabel, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildTeacherSignature();            
            cellFormatDirector.BuildTeacherSignatureCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.TeacherSignature, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildYearFont();            
            cellFormatDirector.BuildYearCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.Year, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildEmployeeInfoMeta();            
            cellFormatDirector.BuildEmployeeInfoMetaCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.EmployeeInfoMeta, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildEmployeeInfo();            
            cellFormatDirector.BuildEmployeeCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
            stylesheetBuilder.AddCellFormat(ExcelCellFormats.EmployeeInfo, cellFormatBuilder.GetCellFormat());

            fontDirector.BuildColumnNumber();
            borderDirector.BuildDefaultBorders();
            cellFormatDirector.BuildColumnNumberCellFormat(stylesheetBuilder.AddFont(fontBuilder.GetFont()), stylesheetBuilder.AddBorder(borderBuilder.GetBorder()));
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