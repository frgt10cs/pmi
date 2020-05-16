using DocumentFormat.OpenXml.Spreadsheet;
using Pmi.Model;

namespace Pmi.Builders
{
    /// <summary>
    /// Строитель stylesheet
    /// </summary>
    public class ExcelStylesheetBuilder
    {
        private ExcelStylesheet stylesheet;                
        private uint fontStartId;
        private uint cellFormatStartId;
        private uint borderStartId;

        /// <param name="fontStartId">Первый свободный идентификатор для шрифта</param>
        /// <param name="cellFormatStartId">Первый свободный идентификатор для формата ячейки</param>
        public ExcelStylesheetBuilder(uint fontStartId, uint cellFormatStartId, uint borderStartId)
        {
            this.fontStartId = fontStartId;
            this.cellFormatStartId = cellFormatStartId;
            this.borderStartId = borderStartId;
            stylesheet = new ExcelStylesheet();            
        }

        /// <summary>
        /// Сбрасывает все значения stylesheet
        /// </summary>
        public void Reset()
        {
            stylesheet.Reset();
        }
        
        /// <summary>
        /// Добавляет шрифт к stylesheet
        /// </summary>
        /// <param name="font"></param>
        /// <returns>Идентификатор добавленного шрифта</returns>
        public uint AddFont(Font font)
        {
            stylesheet.Fonts.Add(font);
            return fontStartId++;
        }

        public uint AddBorder(Border border)
        {
            stylesheet.Borders.Add(border);
            return borderStartId++;
        }

        /// <summary>
        /// Добавляет формат ячейки к stylesheet
        /// </summary>
        /// <param name="cellFormat"></param>
        /// <returns>Идентификатор добавленного формата ячейки</returns>
        public void AddCellFormat(ExcelCellFormats cellFormatType, ExcelCellFormat excelCellFormat)
        {
            excelCellFormat.CellFormatType = cellFormatType;
            excelCellFormat.Id = cellFormatStartId++;
            stylesheet.CellFormats.Add(excelCellFormat);           
        }

        /// <summary>
        /// Добавляет стандартное значение, если один из списков пуст. В случае подгрузки пустых стилей Excel выдаёт ошибку.
        /// </summary>
        private void FillEmpty()
        {
            if (stylesheet.Fonts.Count == 0)
            {
                stylesheet.Fonts.Add(new Font());
            }
            if (stylesheet.CellFormats.Count == 0)
            {
                stylesheet.CellFormats.Add(new ExcelCellFormat());
            }
            if (stylesheet.Borders.Count == 0)
            {
                stylesheet.Borders.Add(new Border());
            }
            if (stylesheet.Fills.Count == 0)
            {
                stylesheet.Fills.Add(new Fill());
            }
        }

        /// <summary>
        /// Возвращает построенный stylesheet
        /// </summary>
        public ExcelStylesheet GetStylesheet()
        {
            FillEmpty();            
            var stylesheetTemp = stylesheet;
            stylesheet = new ExcelStylesheet();            
            return stylesheetTemp;
        }
    }
}
