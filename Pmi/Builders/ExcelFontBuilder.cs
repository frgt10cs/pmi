using DocumentFormat.OpenXml.Spreadsheet;
using Pmi.Model;

namespace Pmi.Builders
{
    /// <summary>
    /// Строитель шрифта
    /// </summary>
    class ExcelFontBuilder
    {
        private Font font;
        public const string defaultFontName = "Times New Roman";        

        public ExcelFontBuilder()
        {
            Reset();
        }

        /// <summary>
        /// Сбрасывает значения шрифта
        /// </summary>
        public void Reset()
        {
            font = new Font();
            SetFontName(defaultFontName);
        }

        /// <summary>
        /// Устаналивает размер шрифта
        /// </summary>
        /// <param name="size"></param>
        public void SetFontSize(double size)
        {
            font.FontSize = new FontSize() { Val = size };
        }

        /// <summary>
        /// Устанавливает тип шрифта
        /// </summary>
        /// <param name="fontName"></param>
        public void SetFontName(string fontName)
        {
            font.FontName = new FontName() { Val = fontName };
        }

        /// <summary>
        /// Устаналивает цвет шрифта
        /// </summary>
        /// <param name="hexValue"></param>
        public void SetColor(string hexValue)
        {
            font.Color = new Color() { Rgb = hexValue };
        }

        /// <summary>
        /// Добалвяет подчеркивание
        /// </summary>
        public void AddUnderline(UnderlineValues underlineValue)
        {
            font.Underline = new Underline() { Val = underlineValue };
        }

        /// <summary>
        /// Добавляет к шрифту жирность
        /// </summary>
        public void AddBold()
        {
            font.Bold = new Bold();
        }

        /// <summary>
        /// Добавляет к шрифту курсив
        /// </summary>
        public void AddItalic()
        {
            font.Italic = new Italic();
        }

        /// <summary>
        /// Возвращает построенный шрифт
        /// </summary>
        /// <returns></returns>
        public Font GetFont()
        {
            Font font = this.font;
            Reset();
            return font;
        }
    }
}
