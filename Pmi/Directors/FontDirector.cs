using Pmi.Builders;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Pmi.Directors
{
    class FontDirector
    {
        private ExcelFontBuilder fontBuilder;
        public ExcelFontBuilder FontBuilder { set => fontBuilder = value; }

        public void BuildUniversityInfoFont()
        {
            fontBuilder.SetFontSize(14);
        }

        public void BuildTitleFont()
        {
            fontBuilder.SetFontSize(16);
            fontBuilder.AddBold();
        }

        public void BuildYearFont()
        {
            fontBuilder.SetFontSize(13);
            fontBuilder.AddUnderline(UnderlineValues.Single);
        }

        public void BuildEmployeeInfoMeta()
        {
            fontBuilder.SetFontSize(8);
        }

        public void BuildEmployeeInfo()
        {
            fontBuilder.SetFontSize(14);
            fontBuilder.AddUnderline(UnderlineValues.Single);
        }

        public void BuildColumnName()
        {
            fontBuilder.SetFontSize(11);
        }

        public void BuildTotal()
        {
            fontBuilder.SetFontSize(11);
            fontBuilder.AddBold();
        }

        public void BuildColumnNumber()
        {
            fontBuilder.SetFontSize(10);
        }

        public void BuildSemesterName()
        {
            fontBuilder.SetFontSize(12);
        }

        public void BuildDisciplineCode()
        {
            fontBuilder.SetFontSize(10);
        }

        public void BuildGroupPlan()
        {
            fontBuilder.SetFontSize(10);
        }

        public void BuildDisciplineName()
        {
            fontBuilder.SetFontSize(10);
        }

        public void BuildColumnTotal()
        {
            fontBuilder.SetFontSize(11);
        }

        public void BuildSemesterTotalName()
        {
            fontBuilder.SetFontSize(12);
        }

        public void BuildApprove()
        {
            fontBuilder.SetFontSize(14);
        }

        public void BuildPosition()
        {
            fontBuilder.SetFontSize(14);
        }

        public void BuildDepatment()
        {
            fontBuilder.SetFontSize(14);
            fontBuilder.AddUnderline(UnderlineValues.Single);
        }

        public void BuildManagerInfo()
        {
            fontBuilder.SetFontSize(12);
            fontBuilder.AddItalic();
        }
        public void BuildManagerInfoMeta()
        {
            fontBuilder.SetFontSize(8);
            fontBuilder.AddItalic();
        }

        public void BuildTeacherSignature()
        {
            fontBuilder.SetFontSize(8);
        }
    }
}
