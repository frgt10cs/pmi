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
            builder.SetType(ExcelCellFormats.UniveristyInfo);
            builder.SetFontId(fontId);
            builder.SetHorizontalAlignment(HorizontalAlignmentValues.Center);
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
    }
}
