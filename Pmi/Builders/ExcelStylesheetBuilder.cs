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

        public ExcelStylesheetBuilder()
        {
            Reset();
        }

        public void Reset()
        {
            stylesheet = new Stylesheet();
            // Создаётся неправильный документ, если в стилях нет хотя бы одного элемента
            // Это эксель так работает
            fonts = new List<Font>() { new Font() };
            cellFormats = new List<CellFormat>() { new CellFormat() };
            borders = new List<Border>() { new Border() };
            fills = new List<Fill>() { new Fill() };
        }

        public uint AddFont(Font font)
        {
            fonts.Add(font);
            // Возвращает id добавленного шрифта
            return Convert.ToUInt32(fonts.Count - 1);
        }

        public void AddCellFormat(CellFormat cellFormat)
        {
            cellFormats.Add(cellFormat);
        }

        public Stylesheet GetStylesheet()
        {
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
