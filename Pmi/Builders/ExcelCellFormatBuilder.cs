using DocumentFormat.OpenXml.Spreadsheet;
using Pmi.Model;

namespace Pmi.Builders
{
    /// <summary>
    /// Строитель формата ячейки
    /// </summary>
    public class ExcelCellFormatBuilder
    {
        private ExcelCellFormat cellFormat;
        private const HorizontalAlignmentValues defaultHorizontalValue = HorizontalAlignmentValues.Left;
        private const VerticalAlignmentValues defaultVerticalValue = VerticalAlignmentValues.Center;

        public ExcelCellFormatBuilder()
        {
            Reset();
        }

        /// <summary>
        /// Сбрасывает значения формата ячейки
        /// </summary>
        public void Reset()
        {
            cellFormat = new ExcelCellFormat
            {
                VerticalAlignment = defaultVerticalValue
            };
        }

        public void SetWrapText(bool wrap)
        {
            cellFormat.Wrap = wrap;
        }

        /// <summary>
        /// Устанавливает горизонтальное выравнивание для ячейки
        /// </summary>
        public void SetHorizontalAlignment(HorizontalAlignmentValues aligment)
        {
            cellFormat.HorizontalAlignment = aligment;
        }

        /// <summary>
        /// Устанавливает вертикальное выравнивание для ячейки
        /// </summary>
        public void SetVerticalAlignment(VerticalAlignmentValues aligment)
        {
            cellFormat.VerticalAlignment = aligment;
        }

        /// <summary>
        /// Устанавливает ссылку на шрифт для ячейки
        /// </summary>
        public void SetFontId(uint fontId)
        {
            cellFormat.FontId = fontId;
        }

        /// <summary>
        /// Устанавливает ссылку на границу ячейки
        /// </summary>
        public void SetBorderId(uint borderId)
        {
            cellFormat.BorderId = borderId;
        }

        /// <summary>
        /// Устанавливает тип для ячейки
        /// </summary>
        public void SetType(ExcelCellFormats type)
        {
            cellFormat.CellFormatType = type;
        }

        /// <summary>
        /// Возвращает построенный формат ячейки
        /// </summary>
        public ExcelCellFormat GetCellFormat()
        {
            var cell = cellFormat;
            Reset();
            return cell;
        }
    }
}
