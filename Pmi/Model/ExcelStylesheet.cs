using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
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
        public List<Font> Fonts { get; set; }        
        public List<ExcelCellFormat> CellFormats { get; set; }        
        public List<Fill> Fills { get; set; }        
        public List<Border> Borders { get; set; }

        public ExcelStylesheet()
        {            
            Reset();            
        }

        /// <summary>
        /// Сбрасывает все значения стилей
        /// </summary>
        public void Reset()
        {            
            Fonts = new List<Font>();
            CellFormats = new List<ExcelCellFormat>();
            Fills = new List<Fill>();
            Borders = new List<Border>();
        }        
    }
}
