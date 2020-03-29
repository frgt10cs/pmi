using DocumentFormat.OpenXml.Spreadsheet;
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
        private CellFormat cellFormat;
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
            cellFormat = new CellFormat();
            cellFormat.Alignment = new Alignment()
            {
                Horizontal = defaultHorizontalValue,
                Vertical = defaultVerticalValue
            };
        }

        /// <summary>
        /// Устанавливает горизонтальное выравнивание для ячейки
        /// </summary>
        /// <param name="aligment"></param>
        public void SetHorizontalAlignment(HorizontalAlignmentValues aligment)
        {
            cellFormat.Alignment.Horizontal = aligment;
        }

        /// <summary>
        /// Устанавливает вертикальное выравнивание для ячейки
        /// </summary>
        /// <param name="aligment"></param>
        public void SetVerticalAlignment(VerticalAlignmentValues aligment)
        {
            cellFormat.Alignment.Vertical = aligment;
        }

        /// <summary>
        /// Добавляет ссылку на шрифт для ячейки
        /// </summary>
        /// <param name="fontId"></param>
        public void AddFontId(uint fontId)
        {
            cellFormat.FontId = fontId;
        }

        /// <summary>
        /// Возвращает построенный формат ячейки
        /// </summary>
        /// <returns></returns>
        public CellFormat GetCellFormat()
        {
            var cell = cellFormat;
            Reset();
            return cell;
        }
    }
}
