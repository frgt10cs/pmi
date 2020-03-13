using DocumentFormat.OpenXml.Spreadsheet;
using Pmi.Service.Interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.Builders
{
    class ExcelFontBuilder:IFontBuilder
    {
        private Font font;
        public const string defaultFontName = "Times New Roman";
        //private const int defaultFontSize = 14;

        public ExcelFontBuilder()
        {
            Reset();
        }

        public void Reset()
        {
            font = new Font();
            AddFontName(defaultFontName);
        }

        public void AddFontSize(int size)
        {
            font.FontSize = new FontSize() { Val = size };
        }

        public void AddFontName(string fontName)
        {
            font.FontName = new FontName() { Val = fontName };
        }

        public void AddColor(string hexValue)
        {
            font.Color = new Color() { Rgb = new DocumentFormat.OpenXml.HexBinaryValue() { Value = hexValue } };
        }

        public void AddUnderline()
        {
            font.Underline = new Underline();
        }

        public void AddBold()
        {
            font.Bold = new Bold();
        }

        public void AddItalic()
        {
            font.Italic = new Italic();
        }

        public Font GetFont()
        {
            Font font = this.font;
            Reset();
            return font;
        }
    }
}
