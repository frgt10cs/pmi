using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Pmi.Model
{
    public class ExcelCellFormat
    {
        /// <summary>
        /// Поле для понимания, к какой ячеки данный формат принадлежит
        /// </summary>
        public ExcelCellFormats CellFormatType;
        public uint Id { get; set; }
        public uint FontId { get; set; }
        public uint BorderId { get; set; }
        public uint FillId { get; set; }
        public HorizontalAlignmentValues HorizontalAlignment { get; set; }
        public VerticalAlignmentValues VerticalAlignment { get; set; }
        public bool Wrap { get; set; } = false;

        public static implicit operator CellFormat(ExcelCellFormat excelCellFormat)
        {
            var format = new CellFormat()
            {
                FontId = excelCellFormat.FontId,
                BorderId = excelCellFormat.BorderId,
                FillId = excelCellFormat.FillId,
                Alignment = new Alignment()
                {
                    Vertical = excelCellFormat.VerticalAlignment,
                    Horizontal = excelCellFormat.HorizontalAlignment,
                    WrapText = excelCellFormat.Wrap
                }
            };

            return format;
        }
    }
}
