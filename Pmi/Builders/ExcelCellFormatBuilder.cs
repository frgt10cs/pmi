using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.Builders
{
    public class ExcelCellFormatBuilder
    {
        private CellFormat cellFormat;
        private const HorizontalAlignmentValues defaultHorizontalValue =
            HorizontalAlignmentValues.Left;
        private const VerticalAlignmentValues defaultVerticalValue =
            VerticalAlignmentValues.Center;

        public ExcelCellFormatBuilder()
        {
            Reset();
        }

        public void Reset()
        {
            cellFormat = new CellFormat();
            cellFormat.Alignment = new Alignment()
            {
                Horizontal = defaultHorizontalValue,
                Vertical = defaultVerticalValue
            };
        }

        public void AddHorizontalAlignment(HorizontalAlignmentValues aligment)
        {
            cellFormat.Alignment.Horizontal = aligment;
        }

        public void AddVerticalAlignment(VerticalAlignmentValues aligment)
        {
            cellFormat.Alignment.Vertical = aligment;
        }

        public void AddFontId(uint fontId)
        {
            cellFormat.FontId = fontId;
        }

        public CellFormat GetCellFormat()
        {
            var cell = cellFormat;
            Reset();
            return cell;
        }
    }
}
