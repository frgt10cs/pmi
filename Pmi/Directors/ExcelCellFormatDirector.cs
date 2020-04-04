using Pmi.Builders;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Pmi.Directors
{
    public class ExcelCellFormatDirector
    {
        private ExcelCellFormatBuilder builder;
        public ExcelCellFormatBuilder CellFormatBuilder { set { builder = value; } }

        public ExcelCellFormatDirector()
        {
            builder = new ExcelCellFormatBuilder();
        }

        public void BuildUniveristyInfoCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Center);
        }

        public void BuildTitleCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Bottom);
        }

        public void BuildEmployeeInfoMetaCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Top);
        }

        public void BuildYearCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Bottom);
        }

        public void BuildEmployeeCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Bottom);
        }

        public void BuildColumnNameCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Center);
            builder.SetWrapText(true);
        }

        public void BuildTotalCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Center);
            builder.SetWrapText(true);
        }

        public void BuildColumnNumberCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Bottom);
        }

        public void BuildDisciplineCodeCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Center);
            builder.SetWrapText(true);
        }

        public void BuildDisciplineNameCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Left);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Center);
            builder.SetWrapText(true);
        }

        public void BuildSemesterNameCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Bottom);
        }

        public void BuildGroupPlanCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Center);
            builder.SetWrapText(true);
        }

        public void BuildSemesterTotalLableCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Bottom);
        }

        public void BuildColumnTotalCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Bottom);
        }

        public void BuildTeacherSignatureCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Top);
        }

        public void BuildApproveCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Left);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Center);
        }

        public void BuildPositionCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Left);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Bottom);
        }

        public void BuildManagerInfoCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Right);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Top);
        }

        public void BuildManagerInfoMetaCellFormat(uint fontId = 0)
        {
            builder.SetFontId(fontId);
            builder.SetVerticalAlignment(VerticalAlignmentValues.Top);
        }
    }
}
