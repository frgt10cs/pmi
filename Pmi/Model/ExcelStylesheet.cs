using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pmi.Model
{
    /// <summary>
    /// Содержит стилевую информацию страницы
    /// </summary>
    public class ExcelStylesheet
    {
        public Dictionary<int, uint> CellFormatIndexes { get; set; }
        public List<Font> Fonts;
        public List<CellFormat> CellFormats;
        public List<Fill> Fills;
        public List<Border> Borders;        

        public ExcelStylesheet()
        {            
            Reset();
        }

        /// <summary>
        /// Сбрасывает все значения стилей
        /// </summary>
        public void Reset()
        {
            CellFormatIndexes = new Dictionary<int, uint>();
            Fonts = new List<Font>();
            CellFormats = new List<CellFormat>();
            Fills = new List<Fill>();
            Borders = new List<Border>();
        }
    }
}
