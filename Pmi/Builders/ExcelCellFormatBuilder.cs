using DocumentFormat.OpenXml.Spreadsheet;
using Pmi.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.Builders
{
    /// <summary>
    /// Строитель формата ячейки
    /// </summary>
    public class ExcelCellFormatBuilder
    {
        private ExcelCellFormat cellFormat;
        private const HorizontalAlignmentValues defaultHorizontalValue =
            HorizontalAlignmentValues.Left;
        private const VerticalAlignmentValues defaultVerticalValue =
            VerticalAlignmentValues.Center;

        public ExcelCellFormatBuilder()
        {
            Reset();
        }

        /// <summary>
        /// Сбрасывает значения формата ячейки
        /// </summary>
        public void Reset()
        {
            cellFormat = new ExcelCellFormat();
            cellFormat.VerticalAlignment = defaultVerticalValue;
        }

        /// <summary>
        /// Устанавливает горизонтальное выравнивание для ячейки
        /// </summary>
        /// <param name="aligment"></param>
        public void SetHorizontalAlignment(HorizontalAlignmentValues aligment)
        {
            cellFormat.HorizontalAlignment = aligment;
        }

        /// <summary>
        /// Устанавливает вертикальное выравнивание для ячейки
        /// </summary>
        /// <param name="aligment"></param>
        public void SetVerticalAlignment(VerticalAlignmentValues aligment)
        {
            cellFormat.VerticalAlignment = aligment;
        }

        /// <summary>
        /// Устанавливает ссылку на шрифт для ячейки
        /// </summary>
        /// <param name="fontId"></param>
        public void SetFontId(uint fontId)
        {
            cellFormat.FontId = fontId;
        }

        /// <summary>
        /// Устанавливает тип для ячейки
        /// </summary>
        /// <param name="type"></param>
        public void SetType(ExcelCellFormats type)
        {
            cellFormat.CellFormatType = type;
        }

        /// <summary>
        /// Устанавливает перенос строки
        /// </summary>
        /// <param name="type"></param>
        public void SetWrapText(bool wrap)
        {
            cellFormat.Wrap = wrap;
        }

        /// <summary>
        /// Возвращает построенный формат ячейки
        /// </summary>
        /// <returns></returns>
        public ExcelCellFormat GetCellFormat()
        {
            var cell = cellFormat;
            Reset();
            return cell;
        }
    }
}
