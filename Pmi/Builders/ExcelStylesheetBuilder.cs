using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.Builders
{
    public class ExcelStylesheetBuilder
    {
        private Stylesheet stylesheet;
        private List<Font> fonts;
        private List<Border> borders;
        private List<Fill> fills;
        private List<CellFormat> cellFormats;
        private uint fontStartId;
        private uint cellFormatStartId;

        /// <summary>
        /// 
        /// </summary>
        /// <param name="fontStartId">Первый свободный идентификатор для шрифта</param>
        /// <param name="cellFormatStartId">Первый свободный идентификатор для формата ячейки</param>
        public ExcelStylesheetBuilder(uint fontStartId, uint cellFormatStartId)
        {
            this.fontStartId = fontStartId;
            this.cellFormatStartId = cellFormatStartId;
            Reset();            
        }

        public void Reset()
        {
            stylesheet = new Stylesheet();            
            fonts = new List<Font>();
            cellFormats = new List<CellFormat>();
            borders = new List<Border>();
            fills = new List<Fill>();
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="font"></param>
        /// <returns>Идентификатор добавленного шрифта</returns>
        public uint AddFont(Font font)
        {
            fonts.Add(font);            
            return fontStartId++;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cellFormat"></param>
        /// <returns>Идентификатор добавленного формата ячейки</returns>
        public uint AddCellFormat(CellFormat cellFormat)
        {
            cellFormats.Add(cellFormat);
            return cellFormatStartId++;
        }

        /// <summary>
        /// Добавляет стандартное значение, если один из списков пуст. В противном случае Excel выдаст ошибку.
        /// </summary>
        private void CheckForEmpty()
        {
            if (fonts.Count == 0)
                fonts.Add(new Font());
            if (cellFormats.Count == 0)
                cellFormats.Add(new CellFormat());
            if (borders.Count == 0)
                borders.Add(new Border());
            if (fills.Count == 0)
                fills.Add(new Fill());
        }

        public Stylesheet GetStylesheet()
        {
            CheckForEmpty();
            stylesheet.Fonts = new Fonts(fonts);
            stylesheet.CellFormats = new CellFormats(cellFormats);
            stylesheet.Borders = new Borders(borders);
            stylesheet.Fills = new Fills(fills);
            var stylesheetTemp = stylesheet;
            Reset();
            return stylesheetTemp;
        }
    }
}
